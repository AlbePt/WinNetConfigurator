using System;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Sockets;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class NetworkService
    {
        public NetworkConfiguration GetActiveConfiguration()
        {
            var adaptersQuery = new SelectQuery("Win32_NetworkAdapter", "NetConnectionStatus IS NOT NULL");
            using (var searcher = new ManagementObjectSearcher(adaptersQuery))
            {
                foreach (ManagementObject adapter in searcher.Get())
                {
                    try
                    {
                        int status = adapter["NetConnectionStatus"] != null ? Convert.ToInt32(adapter["NetConnectionStatus"]) : -1;
                        if (status != 2) continue;

                        string netId = (adapter["NetConnectionID"] as string) ?? string.Empty;
                        string name = (adapter["Name"] as string) ?? string.Empty;
                        string mac = (adapter["MACAddress"] as string) ?? string.Empty;
                        bool isWireless = IsWirelessAdapter(adapter, name, netId);
                        int index = Convert.ToInt32(adapter["Index"]);

                        using (var cfgSearch = new ManagementObjectSearcher(new SelectQuery("Win32_NetworkAdapterConfiguration", $"IPEnabled=true AND Index={index}")))
                        {
                            foreach (ManagementObject cfg in cfgSearch.Get())
                            {
                                var ips = (string[])cfg["IPAddress"];
                                if (ips == null) continue;
                                var ipv4 = ips.FirstOrDefault(IsIPv4);
                                if (string.IsNullOrWhiteSpace(ipv4)) continue;

                                var masks = (string[])cfg["IPSubnet"] ?? Array.Empty<string>();
                                var gateways = (string[])cfg["DefaultIPGateway"] ?? Array.Empty<string>();
                                var dns = (string[])cfg["DNSServerSearchOrder"] ?? Array.Empty<string>();
                                bool dhcp = cfg["DHCPEnabled"] != null && Convert.ToBoolean(cfg["DHCPEnabled"]);

                                return new NetworkConfiguration
                                {
                                    AdapterId = string.IsNullOrWhiteSpace(netId) ? name : netId,
                                    AdapterName = name,
                                    MacAddress = mac,
                                    IpAddress = ipv4,
                                    SubnetMask = masks.FirstOrDefault(IsIPv4) ?? string.Empty,
                                    Gateway = gateways.FirstOrDefault(IsIPv4) ?? string.Empty,
                                    Dns = dns.Where(IsIPv4).ToArray(),
                                    IsDhcpEnabled = dhcp,
                                    IsWireless = isWireless
                                };
                            }
                        }
                    }
                    catch
                    {
                        // ignore broken adapter entries
                    }
                }
            }

            return null;
        }

        public void ApplyConfiguration(string adapterId, string ip, string mask, string gateway, string[] dns)
        {
            try
            {
                ApplyViaWmi(adapterId, ip, mask, gateway, dns);
            }
            catch (Exception wmiError)
            {
                try
                {
                    ApplyViaNetsh(adapterId, ip, mask, gateway, dns);
                }
                catch (Exception netshError)
                {
                    throw new InvalidOperationException(
                        "Не удалось применить настройки сети.\r\n" +
                        "WMI: " + wmiError.Message + "\r\n" +
                        "netsh: " + netshError.Message, netshError);
                }
            }
        }

        static bool IsIPv4(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            if (!IPAddress.TryParse(value, out var ip)) return false;
            return ip.AddressFamily == AddressFamily.InterNetwork;
        }

        void ApplyViaWmi(string adapterId, string ip, string mask, string gateway, string[] dns)
        {
            int index = -1;
            using (var searcher = new ManagementObjectSearcher(new SelectQuery("Win32_NetworkAdapter", $"NetConnectionID='{Escape(adapterId)}'")))
            {
                foreach (ManagementObject adapter in searcher.Get())
                {
                    index = Convert.ToInt32(adapter["Index"]);
                    break;
                }
            }

            if (index < 0)
            {
                using (var searcher = new ManagementObjectSearcher(new SelectQuery("Win32_NetworkAdapter", $"Name='{Escape(adapterId)}'")))
                {
                    foreach (ManagementObject adapter in searcher.Get())
                    {
                        index = Convert.ToInt32(adapter["Index"]);
                        break;
                    }
                }
            }

            if (index < 0)
                throw new InvalidOperationException("Адаптер не найден: " + adapterId);

            using (var cfgSearch = new ManagementObjectSearcher(new SelectQuery("Win32_NetworkAdapterConfiguration", $"IPEnabled=true AND Index={index}")))
            {
                foreach (ManagementObject cfg in cfgSearch.Get())
                {
                    using (var enableStatic = cfg.GetMethodParameters("EnableStatic"))
                    using (var setGateway = cfg.GetMethodParameters("SetGateways"))
                    using (var setDns = cfg.GetMethodParameters("SetDNSServerSearchOrder"))
                    {
                        enableStatic["IPAddress"] = new[] { ip };
                        enableStatic["SubnetMask"] = new[] { mask };
                        cfg.InvokeMethod("EnableStatic", enableStatic, null);

                        setGateway["DefaultIPGateway"] = new[] { gateway };
                        setGateway["GatewayCostMetric"] = new[] { 1 };
                        cfg.InvokeMethod("SetGateways", setGateway, null);

                        setDns["DNSServerSearchOrder"] = dns ?? Array.Empty<string>();
                        cfg.InvokeMethod("SetDNSServerSearchOrder", setDns, null);
                        return;
                    }
                }
            }

            throw new InvalidOperationException("Не удалось получить доступ к параметрам адаптера: " + adapterId);
        }

        void ApplyViaNetsh(string adapterId, string ip, string mask, string gateway, string[] dns)
        {
            string name = adapterId.Replace("\"", "\\\"");
            RunNetsh($"interface ip set address name=\"{name}\" static {ip} {mask} {gateway} 1");
            if (dns != null && dns.Length > 0)
            {
                RunNetsh($"interface ip set dns name=\"{name}\" static {dns[0]} primary");
                for (int i = 1; i < dns.Length; i++)
                    RunNetsh($"interface ip add dns name=\"{name}\" {dns[i]} index={i + 1}");
            }
        }

        void RunNetsh(string args)
        {
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "netsh",
                Arguments = args,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true
            };

            using (var process = System.Diagnostics.Process.Start(psi))
            {
                string stdout = process.StandardOutput.ReadToEnd();
                string stderr = process.StandardError.ReadToEnd();
                process.WaitForExit();
                if (process.ExitCode != 0)
                    throw new InvalidOperationException($"netsh exit {process.ExitCode}: {stdout} {stderr}".Trim());
            }
        }

        string Escape(string value) => value?.Replace("\\", "\\\\").Replace("'", "\\'") ?? string.Empty;

        bool IsWirelessAdapter(ManagementObject adapter, string name, string netId)
        {
            try
            {
                if (adapter["AdapterTypeID"] != null && Convert.ToInt32(adapter["AdapterTypeID"]) == 9)
                    return true;
            }
            catch
            {
                // ignore parse issues
            }

            string product = (adapter["ProductName"] as string) ?? string.Empty;
            string description = (adapter["Description"] as string) ?? string.Empty;

            return ContainsWirelessHint(name)
                || ContainsWirelessHint(netId)
                || ContainsWirelessHint(product)
                || ContainsWirelessHint(description);
        }

        bool ContainsWirelessHint(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            return value.IndexOf("wi-fi", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("wifi", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("wireless", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("802.11", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
