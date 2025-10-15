using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Sockets;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class NetworkService
    {
        public IEnumerable<AdapterInfo> ListAdapters(bool includeVirtual)
        {
            var result = new List<AdapterInfo>();

            var qs = new SelectQuery("Win32_NetworkAdapter", "NetConnectionStatus IS NOT NULL");
            using (var s = new ManagementObjectSearcher(qs))
            {
                foreach (ManagementObject na in s.Get())
                {
                    try
                    {
                        string pnp = (na["PNPDeviceID"] as string) ?? "";
                        string manuf = (na["Manufacturer"] as string) ?? "";
                        bool looksVirtual =
                            pnp.IndexOf("VMS_", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            pnp.IndexOf("TAP", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            pnp.IndexOf("ROOT\\", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            manuf.IndexOf("Virtual", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            manuf.IndexOf("VPN", StringComparison.OrdinalIgnoreCase) >= 0;

                        if (!includeVirtual && looksVirtual) continue;

                        int index = Convert.ToInt32(na["Index"]);
                        string netId = (na["NetConnectionID"] as string) ?? "";
                        string name = (na["Name"] as string) ?? "";
                        string mac = (na["MACAddress"] as string) ?? "";
                        string type = DetectType(na);

                        string ipv4 = "";
                        using (var cfgSearch = new ManagementObjectSearcher(
                            new SelectQuery("Win32_NetworkAdapterConfiguration", $"IPEnabled=true AND Index={index}")))
                        {
                            foreach (ManagementObject cfg in cfgSearch.Get())
                            {
                                var ips = (string[])cfg["IPAddress"];
                                if (ips != null)
                                {
                                    ipv4 = ips.FirstOrDefault(x =>
                                        IPAddress.TryParse(x, out var ip) &&
                                        ip.AddressFamily == AddressFamily.InterNetwork) ?? "";
                                }
                                break;
                            }
                        }

                        result.Add(new AdapterInfo
                        {
                            NetConnectionId = netId,
                            Name = name,
                            MacAddress = mac,
                            IPv4Address = ipv4,
                            Type = type
                        });
                    }
                    catch
                    {
                        // игнорируем проблемные адаптеры
                    }
                }
            }

            return result
                .OrderByDescending(a => a.Type == "Ethernet")
                .ThenBy(a => a.NetConnectionId ?? a.Name)
                .ToList();
        }

        private string DetectType(ManagementObject na)
        {
            var adapterType = (na["AdapterType"] as string) ?? "";
            var name = (na["Name"] as string) ?? "";
            var netId = (na["NetConnectionID"] as string) ?? "";
            string joined = string.Join(" ", adapterType, name, netId).ToLowerInvariant();
            if (joined.Contains("wireless") || joined.Contains("wi-fi") || joined.Contains("wifi")) return "Wi-Fi";
            if (joined.Contains("802.11")) return "Wi-Fi";
            return "Ethernet";
        }

        public void ApplyIPv4(string adapterId, string ip, string mask, string gateway, string[] dns)
        {
            try
            {
                ApplyIPv4_Wmi(adapterId, ip, mask, gateway, dns);
                return;
            }
            catch (Exception wmiEx)
            {
                try
                {
                    ApplyIPv4_Netsh(adapterId, ip, mask, gateway, dns);
                    return;
                }
                catch (Exception netshEx)
                {
                    throw new InvalidOperationException(
                        "WMI и netsh не смогли применить IPv4.\r\nWMI: " + wmiEx.Message + "\r\nnetsh: " + netshEx.Message, netshEx);
                }
            }
        }

        private void ApplyIPv4_Wmi(string adapterId, string ip, string mask, string gateway, string[] dns)
        {
            int index = -1;
            using (var s = new ManagementObjectSearcher(new SelectQuery("Win32_NetworkAdapter", $"NetConnectionID='{EscapeWmi(adapterId)}'")))
            {
                foreach (ManagementObject na in s.Get())
                {
                    index = Convert.ToInt32(na["Index"]);
                    break;
                }
            }
            if (index < 0) throw new InvalidOperationException("Адаптер не найден: " + adapterId);

            using (var cfgSearch = new ManagementObjectSearcher(new SelectQuery("Win32_NetworkAdapterConfiguration", $"IPEnabled=true AND Index={index}")))
            {
                foreach (ManagementObject cfg in cfgSearch.Get())
                {
                    using (var enableStatic = cfg.GetMethodParameters("EnableStatic"))
                    using (var setGw = cfg.GetMethodParameters("SetGateways"))
                    using (var setDns = cfg.GetMethodParameters("SetDNSServerSearchOrder"))
                    {
                        enableStatic["IPAddress"] = new[] { ip };
                        enableStatic["SubnetMask"] = new[] { mask };
                        var r1 = cfg.InvokeMethod("EnableStatic", enableStatic, null);

                        setGw["DefaultIPGateway"] = new[] { gateway };
                        setGw["GatewayCostMetric"] = new[] { 1 };
                        var r2 = cfg.InvokeMethod("SetGateways", setGw, null);

                        setDns["DNSServerSearchOrder"] = dns;
                        var r3 = cfg.InvokeMethod("SetDNSServerSearchOrder", setDns, null);

                        return;
                    }
                }
            }
            throw new InvalidOperationException("Не найден IPEnabled конфиг для адаптера: " + adapterId);
        }

        private void ApplyIPv4_Netsh(string adapterId, string ip, string mask, string gateway, string[] dns)
        {
            string name = adapterId.Replace("\"", "\\\"");
            RunNetsh($@"interface ip set address name=""{name}"" static {ip} {mask} {gateway} 1");
            if (dns != null && dns.Length > 0)
            {
                RunNetsh($@"interface ip set dns name=""{name}"" static {dns[0]} primary");
                for (int i = 1; i < dns.Length; i++)
                    RunNetsh($@"interface ip add dns name=""{name}"" {dns[i]} index={i + 1}");
            }
        }

        private void RunNetsh(string args)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "netsh",
                Arguments = args,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using (var p = Process.Start(psi))
            {
                string stdout = p.StandardOutput.ReadToEnd();
                string stderr = p.StandardError.ReadToEnd();
                p.WaitForExit();
                if (p.ExitCode != 0)
                    throw new InvalidOperationException($"netsh exit {p.ExitCode}: {stdout} {stderr}".Trim());
            }
        }

        private string EscapeWmi(string s) => s?.Replace("\\", "\\\\").Replace("'", "\\'") ?? "";
    }
}
