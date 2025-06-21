using System;
using System.Collections.Generic;
using System.Management;
using System.Net.NetworkInformation;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;

namespace HardwareChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Windows 电脑硬件检测工具");
            Console.WriteLine("========================");
            Console.WriteLine();

            try
            {
                CheckMainboard();
                CheckFans();
                CheckMemory();
                CheckStorage();
                CheckTemperatures();
                CheckGPU();
                CheckNetwork();
                CheckDrivers();
                CheckPower();
                CheckOS();
                CheckCPU();
                CheckLaptopSpecific();
                CheckOtherDevices();

                Console.WriteLine();
                Console.WriteLine("检测完成!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("检测过程中发生错误: " + ex.Message);
            }

            Console.WriteLine();
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }

        static void PrintCheckResult(string category, string item, string status, string message)
        {
            Console.ForegroundColor = status == "正常" ? ConsoleColor.Green : 
                                    status == "警告" ? ConsoleColor.Yellow : 
                                    status == "错误" ? ConsoleColor.Red : ConsoleColor.White;
            
            Console.WriteLine("[" + category + "] " + item + ": " + message + " (" + status + ")");
            Console.ResetColor();
        }

        static void CheckMainboard()
        {
            Console.WriteLine();
            Console.WriteLine("=== 主板检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BIOS");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string biosVersion = obj["Version"] != null ? obj["Version"].ToString() : "未知";
                    PrintCheckResult("主板", "BIOS版本", 
                        IsBiosOld(biosVersion) ? "警告" : "正常", 
                        biosVersion);
                }

                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard");
                foreach (ManagementObject obj in searcher.Get())
                {
                    PrintCheckResult("主板", "主板信息", "正常", 
                        "制造商: " + obj["Manufacturer"] + ", 型号: " + obj["Product"]);
                }

                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_VoltageProbe");
                foreach (ManagementObject obj in searcher.Get())
                {
                    PrintCheckResult("主板", "电压检测", "正常", 
                        obj["Name"] + ": " + obj["CurrentReading"] + "V");
                }

                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity");
                int connectedDevices = 0;
                foreach (ManagementObject obj in searcher.Get())
                {
                    if (obj["Status"] != null && obj["Status"].ToString() == "OK") connectedDevices++;
                }
                PrintCheckResult("主板", "IO设备连接", "正常", 
                    "检测到 " + connectedDevices + " 个正常连接的设备");
            }
            catch (Exception ex)
            {
                PrintCheckResult("主板", "检查错误", "错误", ex.Message);
            }
        }

        static bool IsBiosOld(string version)
        {
            if (string.IsNullOrEmpty(version)) return false;
            
            return version.Contains("2018") || version.Contains("2017") || 
                   version.Contains("2016") || version.Contains("2015");
        }

        static void CheckFans()
        {
            Console.WriteLine();
            Console.WriteLine("=== 风扇检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Fan");
                var fans = searcher.Get();
                
                if (fans.Count == 0)
                {
                    PrintCheckResult("风扇", "检测", "警告", "未检测到风扇信息");
                    return;
                }

                foreach (ManagementObject fan in fans)
                {
                    string status = fan["Status"] != null && fan["Status"].ToString() == "OK" ? "正常" : "警告";
                    PrintCheckResult("风扇", fan["Name"] != null ? fan["Name"].ToString() : "未知风扇", status, 
                        "转速: " + fan["DesiredSpeed"] + " RPM");
                }
            }
            catch (Exception ex)
            {
                PrintCheckResult("风扇", "检查错误", "错误", ex.Message);
            }
        }

        static void CheckMemory()
        {
            Console.WriteLine();
            Console.WriteLine("=== 内存检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                foreach (ManagementObject obj in searcher.Get())
                {
                    double totalMemory = Convert.ToDouble(obj["TotalPhysicalMemory"]) / 1024 / 1024 / 1024;
                    PrintCheckResult("内存", "总内存", "正常", string.Format("{0:F2} GB", totalMemory));
                }

                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string manufacturer = obj["Manufacturer"] != null ? obj["Manufacturer"].ToString() : "未知";
                    double capacity = Convert.ToDouble(obj["Capacity"]) / 1024 / 1024 / 1024;
                    int speed = Convert.ToInt32(obj["Speed"]);
                    
                    PrintCheckResult("内存", "内存模块", "正常", 
                        manufacturer + ", " + string.Format("{0:F2}", capacity) + " GB, " + speed + " MHz");
                }

                PerformMemoryBenchmark();
            }
            catch (Exception ex)
            {
                PrintCheckResult("内存", "检查错误", "错误", ex.Message);
            }
        }

        static void PerformMemoryBenchmark()
        {
            try
            {
                Stopwatch sw = Stopwatch.StartNew();
                byte[] testArray = new byte[100000000];
                new Random().NextBytes(testArray);
                
                for (int i = 0; i < testArray.Length; i += 4096)
                {
                    testArray[i] = (byte)(testArray[i] * 0.5);
                }
                
                sw.Stop();
                double speed = 100.0 / (sw.ElapsedMilliseconds / 1000.0);
                
                string status = speed > 5000 ? "正常" : "警告";
                PrintCheckResult("内存", "基准测试", status, "速度: " + string.Format("{0:F2}", speed) + " MB/s");
            }
            catch (Exception ex)
            {
                PrintCheckResult("内存", "基准测试错误", "错误", ex.Message);
            }
        }

        static void CheckStorage()
        {
            Console.WriteLine();
            Console.WriteLine("=== 存储设备检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
                foreach (ManagementObject disk in searcher.Get())
                {
                    string model = disk["Model"] != null ? disk["Model"].ToString().Trim() : "未知磁盘";
                    double size = Convert.ToDouble(disk["Size"]) / 1024 / 1024 / 1024;
                    
                    PrintCheckResult("存储", "磁盘", "正常", model + ", " + string.Format("{0:F2}", size) + " GB");

                    string smartStatus = disk["Status"] != null && disk["Status"].ToString() == "OK" ? "正常" : "警告";
                    PrintCheckResult("存储", "SMART状态", smartStatus, model);
                }

                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PerfFormattedData_PerfDisk_PhysicalDisk");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string name = obj["Name"] != null ? obj["Name"].ToString() : "";
                    if (name == "_Total") continue;
                    
                    int busyPercent = Convert.ToInt32(obj["PercentDiskTime"]);
                    string status = busyPercent > 90 ? "警告" : "正常";
                    PrintCheckResult("存储", "IO占用", status, name + ": " + busyPercent + "%");
                }
            }
            catch (Exception ex)
            {
                PrintCheckResult("存储", "检查错误", "错误", ex.Message);
            }
        }

        [DllImport("kernel32.dll")]
        static extern uint GetTickCount();

        static void CheckTemperatures()
        {
            Console.WriteLine();
            Console.WriteLine("=== 温度检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_TemperatureProbe");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string name = obj["Name"] != null ? obj["Name"].ToString() : "温度传感器";
                    int? reading = obj["CurrentReading"] as int?;
                    
                    if (reading.HasValue)
                    {
                        string status = reading > 80 ? "警告" : "正常";
                        PrintCheckResult("温度", name, status, reading + "°C");
                    }
                    else
                    {
                        PrintCheckResult("温度", name, "警告", "无法获取读数");
                    }
                }

                searcher = new ManagementObjectSearcher(@"root\WMI", "SELECT * FROM MSAcpi_ThermalZoneTemperature");
                foreach (ManagementObject obj in searcher.Get())
                {
                    double tempKelvin = Convert.ToDouble(obj["CurrentTemperature"].ToString());
                    double tempCelsius = tempKelvin / 10 - 273.15;
                    string status = tempCelsius > 80 ? "警告" : "正常";
                    PrintCheckResult("温度", "CPU温度", status, string.Format("{0:F1}", tempCelsius) + "°C");
                }
            }
            catch (Exception ex)
            {
                PrintCheckResult("温度", "检查错误", "错误", ex.Message);
            }
        }

        static void CheckGPU()
        {
            Console.WriteLine();
            Console.WriteLine("=== 显卡检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController");
                foreach (ManagementObject gpu in searcher.Get())
                {
                    string name = gpu["Name"] != null ? gpu["Name"].ToString() : "未知显卡";
                    string status = "正常";

                    if (name.Contains("NVIDIA") && gpu["AdapterRAM"] != null)
                    {
                        long ram = Convert.ToInt64(gpu["AdapterRAM"]);
                        if (ram < 1024 * 1024 * 1024)
                        {
                            status = "警告";
                            name += " (可能是假卡)";
                        }
                    }

                    PrintCheckResult("显卡", "检测", status, name);

                    string driverStatus = gpu["Status"] != null && gpu["Status"].ToString() == "OK" ? "正常" : "警告";
                    PrintCheckResult("显卡", "驱动状态", driverStatus, 
                        "版本: " + gpu["DriverVersion"]);

                    PrintCheckResult("显卡", "PCIe速率", "信息", "需要专用工具检测");
                }
            }
            catch (Exception ex)
            {
                PrintCheckResult("显卡", "检查错误", "错误", ex.Message);
            }
        }

        static void CheckNetwork()
        {
            Console.WriteLine();
            Console.WriteLine("=== 网络检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled=true");
                foreach (ManagementObject adapter in searcher.Get())
                {
                    string name = adapter["Name"] != null ? adapter["Name"].ToString() : "未知适配器";
                    string status = adapter["NetConnectionStatus"] != null && adapter["NetConnectionStatus"].ToString() == "2" ? "正常" : "警告";
                    
                    PrintCheckResult("网络", "适配器", status, name);
                }

                bool networkAvailable = NetworkInterface.GetIsNetworkAvailable();
                PrintCheckResult("网络", "网络连接", networkAvailable ? "正常" : "错误", 
                    networkAvailable ? "已连接" : "未连接");

                try
                {
                    IPHostEntry entry = Dns.GetHostEntry("www.microsoft.com");
                    PrintCheckResult("网络", "DNS解析", "正常", 
                        "成功解析: " + entry.HostName);
                }
                catch
                {
                    PrintCheckResult("网络", "DNS解析", "错误", "解析失败");
                }

                string hostsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "drivers/etc/hosts");
                if (File.Exists(hostsPath))
                {
                    var hostsContent = File.ReadAllText(hostsPath);
                    bool hasSuspiciousEntries = hostsContent.Contains("127.0.0.1 microsoft.com") || 
                                              hostsContent.Contains("127.0.0.1 google.com");
                    PrintCheckResult("网络", "hosts文件", 
                        hasSuspiciousEntries ? "警告" : "正常", 
                        hasSuspiciousEntries ? "检测到可疑条目" : "正常");
                }

                searcher = new ManagementObjectSearcher(@"root\SecurityCenter2", "SELECT * FROM FirewallProduct");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string firewallStatus = obj["productState"] != null && obj["productState"].ToString() == "266240" ? "启用" : "禁用";
                    PrintCheckResult("网络", "防火墙", 
                        firewallStatus == "启用" ? "正常" : "警告", 
                        "状态: " + firewallStatus);
                }
            }
            catch (Exception ex)
            {
                PrintCheckResult("网络", "检查错误", "错误", ex.Message);
            }
        }

        static void CheckDrivers()
        {
            Console.WriteLine();
            Console.WriteLine("=== 驱动检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode <> 0");
                var problemDevices = searcher.Get();
                
                if (problemDevices.Count == 0)
                {
                    PrintCheckResult("驱动", "状态", "正常", "未检测到有问题的驱动");
                }
                else
                {
                    foreach (ManagementObject device in problemDevices)
                    {
                        string name = device["Name"] != null ? device["Name"].ToString() : "未知设备";
                        int errorCode = Convert.ToInt32(device["ConfigManagerErrorCode"]);
                        PrintCheckResult("驱动", "问题设备", "错误", name + " (错误代码: " + errorCode + ")");
                    }
                }
            }
            catch (Exception ex)
            {
                PrintCheckResult("驱动", "检查错误", "错误", ex.Message);
            }
        }

        static void CheckPower()
        {
            Console.WriteLine();
            Console.WriteLine("=== 电源检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Battery");
                var batteries = searcher.Get();
                
                if (batteries.Count == 0)
                {
                    PrintCheckResult("电源", "类型", "信息", "台式机或未检测到电池");
                }
                else
                {
                    foreach (ManagementObject battery in batteries)
                    {
                        string status = battery["BatteryStatus"] != null && battery["BatteryStatus"].ToString() == "2" ? "充电中" : "使用电池";
                        int? chargeRemaining = battery["EstimatedChargeRemaining"] as int?;
                        
                        string chargeStatus = chargeRemaining.HasValue && chargeRemaining < 20 ? "警告" : "正常";
                        PrintCheckResult("电源", "电池状态", chargeStatus, 
                            status + ", 剩余电量: " + chargeRemaining + "%");
                    }
                }

                PrintCheckResult("电源", "输出电压", "信息", "需要专用工具检测");
            }
            catch (Exception ex)
            {
                PrintCheckResult("电源", "检查错误", "错误", ex.Message);
            }
        }

        static void CheckOS()
        {
            Console.WriteLine();
            Console.WriteLine("=== 操作系统检查 ===");
            
            try
            {
                var osInfo = Environment.OSVersion;
                PrintCheckResult("OS", "版本", "信息", 
                    osInfo.VersionString + ", " + GetOSFriendlyName());

                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    string productName = key != null && key.GetValue("ProductName") != null ? key.GetValue("ProductName").ToString() : "未知";
                    string editionId = key != null && key.GetValue("EditionID") != null ? key.GetValue("EditionID").ToString() : "未知";
                    PrintCheckResult("OS", "授权", "信息", productName + " (" + editionId + ")");
                }

                CheckBlueScreenLogs();
            }
            catch (Exception ex)
            {
                PrintCheckResult("OS", "检查错误", "错误", ex.Message);
            }
        }

        static string GetOSFriendlyName()
        {
            using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
            {
                return key != null && key.GetValue("ProductName") != null ? key.GetValue("ProductName").ToString() : "未知Windows版本";
            }
        }

        static void CheckBlueScreenLogs()
        {
            try
            {
                var searcher = new ManagementObjectSearcher(
                    "SELECT * FROM Win32_NTLogEvent WHERE LogFile='System' AND EventCode=41");
                
                int crashCount = 0;
                DateTime lastCrash = DateTime.MinValue;
                
                foreach (ManagementObject log in searcher.Get())
                {
                    crashCount++;
                    DateTime time = ManagementDateTimeConverter.ToDateTime(log["TimeGenerated"].ToString());
                    if (time > lastCrash) lastCrash = time;
                }

                string status = crashCount == 0 ? "正常" : "警告";
                string message = crashCount == 0 ? "最近一个月无蓝屏记录" : 
                    "最近一个月蓝屏次数: " + crashCount + ", 最后一次: " + lastCrash.ToString("yyyy-MM-dd");
                
                PrintCheckResult("OS", "系统稳定性", status, message);
            }
            catch (Exception ex)
            {
                PrintCheckResult("OS", "蓝屏日志检查", "错误", ex.Message);
            }
        }

        static void CheckCPU()
        {
            Console.WriteLine();
            Console.WriteLine("=== CPU检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                foreach (ManagementObject cpu in searcher.Get())
                {
                    string name = cpu["Name"] != null ? cpu["Name"].ToString() : "未知CPU";
                    int cores = Convert.ToInt32(cpu["NumberOfCores"]);
                    int threads = Convert.ToInt32(cpu["NumberOfLogicalProcessors"]);
                    int currentClock = Convert.ToInt32(cpu["CurrentClockSpeed"]);
                    int maxClock = Convert.ToInt32(cpu["MaxClockSpeed"]);
                    
                    PrintCheckResult("CPU", "信息", "正常", 
                        name + ", " + cores + "核/" + threads + "线程, 频率: " + currentClock + "/" + maxClock + " MHz");

                    PerformCpuBenchmark();
                }
            }
            catch (Exception ex)
            {
                PrintCheckResult("CPU", "检查错误", "错误", ex.Message);
            }
        }

        static void PerformCpuBenchmark()
        {
            try
            {
                Stopwatch sw = Stopwatch.StartNew();
                double sum = 0;
                for (int i = 0; i < 100000000; i++)
                {
                    sum += Math.Sqrt(i);
                }
                sw.Stop();
                
                double score = 100000000 / (sw.ElapsedMilliseconds / 1000.0);
                string status = score > 50000000 ? "正常" : "警告";
                
                PrintCheckResult("CPU", "基准测试", status, 
                    "得分: " + string.Format("{0:F0}", score) + " ops/s (耗时: " + sw.ElapsedMilliseconds + "ms)");
            }
            catch (Exception ex)
            {
                PrintCheckResult("CPU", "基准测试", "错误", ex.Message);
            }
        }

        static void CheckLaptopSpecific()
        {
            Console.WriteLine();
            Console.WriteLine("=== 笔记本专用检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Battery");
                var batteries = searcher.Get();
                
                if (batteries.Count == 0)
                {
                    PrintCheckResult("笔记本", "检测", "信息", "未检测到电池，可能是台式机");
                    return;
                }

                PrintCheckResult("笔记本", "网卡温度", "信息", "需要专用工具检测");
            }
            catch (Exception ex)
            {
                PrintCheckResult("笔记本", "检查错误", "错误", ex.Message);
            }
        }

        static void CheckOtherDevices()
        {
            Console.WriteLine();
            Console.WriteLine("=== 其他设备检查 ===");
            
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_USBControllerDevice");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string dependent = obj["Dependent"] != null ? obj["Dependent"].ToString() : "";
                    string deviceName = dependent.Split(new[] { '=' }, 2)[1].Trim('"');
                    PrintCheckResult("USB设备", "检测", "正常", deviceName);
                }

                searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE PNPClass='System'");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string name = obj["Name"] != null ? obj["Name"].ToString() : "未知设备";
                    PrintCheckResult("PCIe设备", "检测", "正常", name);
                }

                PrintCheckResult("设备", "协商速率", "信息", "需要专用工具检测");
            }
            catch (Exception ex)
            {
                PrintCheckResult("设备", "检查错误", "错误", ex.Message);
            }
        }
    }
}
