import os
import sys
import wmi
import psutil
import platform
import GPUtil
import speedtest
import win32api
import win32con
import win32file
import pythoncom
import subprocess
from datetime import datetime
from tabulate import tabulate

class SystemChecker:
    def __init__(self):
        pythoncom.CoInitialize()
        self.c = wmi.WMI()
        self.results = []
        self.errors = []
        
    def check_mainboard(self):
        """检查主板相关信息"""
        try:
            # 检查BIOS版本
            bios = self.c.Win32_BIOS()[0]
            self.results.append(("主板/BIOS", f"BIOS版本: {bios.Version}", "信息"))
            
            # 检查主板信息
            mb = self.c.Win32_BaseBoard()[0]
            self.results.append(("主板", f"制造商: {mb.Manufacturer}, 型号: {mb.Product}", "信息"))
            
        except Exception as e:
            self.errors.append(("主板检查", f"错误: {str(e)}", "错误"))
    
    def check_fans(self):
        """检查风扇状态"""
        try:
            fans = self.c.Win32_Fan()
            if not fans:
                self.results.append(("风扇", "未检测到风扇信息", "警告"))
                return
                
            for fan in fans:
                status = "正常" if fan.Status == "OK" else "异常"
                self.results.append(("风扇", f"设备: {fan.Name}, 状态: {status}", status))
                
        except Exception as e:
            self.errors.append(("风扇检查", f"错误: {str(e)}", "错误"))
    
    def check_memory(self):
        """检查内存状态"""
        try:
            # 内存总量
            mem = psutil.virtual_memory()
            self.results.append(("内存", f"总内存: {mem.total/1024/1024/1024:.2f} GB", "信息"))
            
            # 内存频率和时序需要通过其他方式获取
            # 这里只是示例
            mem_modules = self.c.Win32_PhysicalMemory()
            for i, module in enumerate(mem_modules):
                self.results.append((f"内存模块{i+1}", 
                                   f"容量: {int(module.Capacity)/1024/1024/1024:.2f} GB, 速度: {module.Speed} MHz", 
                                   "信息"))
                
        except Exception as e:
            self.errors.append(("内存检查", f"错误: {str(e)}", "错误"))
    
    def check_storage(self):
        """检查存储设备"""
        try:
            disks = self.c.Win32_DiskDrive()
            for disk in disks:
                # 检查SMART状态
                model = disk.Model.strip()
                size = int(disk.Size)/1024/1024/1024 if disk.Size else 0
                self.results.append(("存储设备", f"型号: {model}, 大小: {size:.2f} GB", "信息"))
                
                # 检查IO占用
                for partition in self.c.Win32_DiskPartition(DeviceID=disk.DeviceID):
                    try:
                        disk_usage = psutil.disk_usage(partition.DeviceID.replace("\\", ""))
                        self.results.append(("存储IO", f"分区 {partition.DeviceID} 使用率: {disk_usage.percent}%", 
                                           "警告" if disk_usage.percent > 90 else "正常"))
                    except:
                        pass
                        
        except Exception as e:
            self.errors.append(("存储检查", f"错误: {str(e)}", "错误"))
    
    def check_temperatures(self):
        """检查温度"""
        try:
            # 需要安装OpenHardwareMonitor或其他温度监控工具
            # 这里只是示例
            temps = self.c.Win32_TemperatureProbe()
            for temp in temps:
                self.results.append(("温度传感器", f"{temp.Name}: {temp.CurrentReading}°C", 
                                   "警告" if temp.CurrentReading > 80 else "正常"))
                                   
        except Exception as e:
            self.errors.append(("温度检查", f"错误: {str(e)}", "错误"))
    
    def check_gpu(self):
        """检查显卡状态"""
        try:
            gpus = GPUtil.getGPUs()
            if not gpus:
                self.results.append(("显卡", "未检测到独立显卡", "信息"))
                return
                
            for gpu in gpus:
                status = "正常"
                if gpu.temperature > 85:
                    status = "警告"
                if "fake" in gpu.name.lower():
                    status = "错误 - 可能是假卡"
                    
                self.results.append(("显卡", 
                                    f"型号: {gpu.name}, 温度: {gpu.temperature}°C, 负载: {gpu.load*100:.1f}%", 
                                    status))
                                    
        except Exception as e:
            self.errors.append(("显卡检查", f"错误: {str(e)}", "错误"))
    
    def check_network(self):
        """检查网络状态"""
        try:
            # 检查网络适配器
            adapters = self.c.Win32_NetworkAdapter(NetEnabled=True)
            for adapter in adapters:
                if adapter.NetConnectionStatus == 2:  # 已连接
                    self.results.append(("网络适配器", f"{adapter.Name} - 已连接", "正常"))
                    
            # 检查网络连接
            net_stats = psutil.net_io_counters()
            self.results.append(("网络IO", f"发送: {net_stats.bytes_sent/1024/1024:.2f} MB, 接收: {net_stats.bytes_recv/1024/1024:.2f} MB", "信息"))
            
            # 检查DNS
            try:
                dns_result = subprocess.check_output("nslookup google.com", shell=True).decode()
                if "Non-authoritative answer" in dns_result:
                    self.results.append(("DNS", "DNS解析正常", "正常"))
                else:
                    self.results.append(("DNS", "DNS解析可能有问题", "警告"))
            except:
                self.results.append(("DNS", "DNS解析测试失败", "警告"))
                
        except Exception as e:
            self.errors.append(("网络检查", f"错误: {str(e)}", "错误"))
    
    def check_drivers(self):
        """检查驱动程序状态"""
        try:
            problem_devices = self.c.Win32_PnPEntity(ConfigManagerErrorCode__ne=0)
            for device in problem_devices:
                self.errors.append(("驱动问题", 
                                  f"设备: {device.Name}, 错误代码: {device.ConfigManagerErrorCode}", 
                                  "错误"))
                                  
            self.results.append(("驱动检查", 
                               f"发现 {len(problem_devices)} 个有问题的设备" if problem_devices else "所有驱动正常", 
                               "错误" if problem_devices else "正常"))
                               
        except Exception as e:
            self.errors.append(("驱动检查", f"错误: {str(e)}", "错误"))
    
    def check_power(self):
        """检查电源状态"""
        try:
            # 检查电池状态（笔记本）
            batteries = self.c.Win32_Battery()
            if batteries:
                for battery in batteries:
                    status = "充电中" if battery.BatteryStatus == 2 else "使用电池"
                    self.results.append(("电源", 
                                       f"电池状态: {status}, 剩余电量: {battery.EstimatedChargeRemaining}%", 
                                       "警告" if battery.EstimatedChargeRemaining < 20 else "正常"))
            else:
                self.results.append(("电源", "台式机或未检测到电池", "信息"))
                
        except Exception as e:
            self.errors.append(("电源检查", f"错误: {str(e)}", "错误"))
    
    def check_os(self):
        """检查操作系统状态"""
        try:
            # 系统信息
            os_info = platform.uname()
            self.results.append(("操作系统", 
                               f"系统: {os_info.system}, 版本: {os_info.version}, 发行版: {os_info.release}", 
                               "信息"))
                               
            # 检查蓝屏日志
            try:
                event_log = subprocess.check_output(
                    'wevtutil qe System "/q:*[System[(EventID=41)]]" /f:text', 
                    shell=True).decode()
                if "Kernel-Power" in event_log:
                    self.errors.append(("系统日志", "检测到近期蓝屏记录", "错误"))
            except:
                pass
                
        except Exception as e:
            self.errors.append(("操作系统检查", f"错误: {str(e)}", "错误"))
    
    def check_cpu(self):
        """检查CPU状态"""
        try:
            # CPU信息
            cpu_info = self.c.Win32_Processor()[0]
            self.results.append(("CPU", 
                               f"型号: {cpu_info.Name}, 核心数: {cpu_info.NumberOfCores}, 线程数: {cpu_info.NumberOfLogicalProcessors}", 
                               "信息"))
                               
            # CPU频率
            self.results.append(("CPU频率", 
                               f"当前频率: {cpu_info.CurrentClockSpeed} MHz, 最大频率: {cpu_info.MaxClockSpeed} MHz", 
                               "正常"))
                               
            # CPU负载
            cpu_load = psutil.cpu_percent(interval=1)
            self.results.append(("CPU负载", 
                               f"当前负载: {cpu_load}%", 
                               "警告" if cpu_load > 90 else "正常"))
                               
        except Exception as e:
            self.errors.append(("CPU检查", f"错误: {str(e)}", "错误"))
    
    def run_all_checks(self):
        """运行所有检查"""
        self.results.append(("系统检查", f"开始于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "信息"))
        
        # 运行各个检查模块
        self.check_mainboard()
        self.check_fans()
        self.check_memory()
        self.check_storage()
        self.check_temperatures()
        self.check_gpu()
        self.check_network()
        self.check_drivers()
        self.check_power()
        self.check_os()
        self.check_cpu()
        
        self.results.append(("系统检查", f"完成于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "信息"))
        
    def print_report(self):
        """打印检查报告"""
        print("\n" + "="*50)
        print(" Windows 电脑问题检查报告")
        print("="*50 + "\n")
        
        # 打印正常结果
        print(tabulate(self.results, headers=["检查项", "结果", "状态"], tablefmt="grid"))
        
        # 打印错误信息
        if self.errors:
            print("\n" + "!"*50)
            print(" 发现的问题和错误:")
            print("!"*50 + "\n")
            print(tabulate(self.errors, headers=["检查项", "问题", "严重性"], tablefmt="grid"))
        else:
            print("\n没有发现严重问题!")
            
        print("\n提示: 某些检查可能需要管理员权限才能获取完整信息")
        print("对于标记为警告或错误的问题，建议进一步排查\n")

if __name__ == "__main__":
    print("正在检查系统...")
    checker = SystemChecker()
    checker.run_all_checks()
    checker.print_report()
    
    # 保持窗口打开
    if os.name == 'nt':
        os.system("pause")
