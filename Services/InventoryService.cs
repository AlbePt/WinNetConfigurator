using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;

namespace WinNetConfigurator.Services
{
    public class InventoryService
    {
        public (string ip, string mask, string gateway, string[] dns) ReadCurrentIPv4(string adapterNameHint = null)
        {
            try
            {
                string q = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True";
                using (var s = new ManagementObjectSearcher(q))
                {
                    foreach (ManagementObject conf in s.Get())
                    {
                        string desc = (conf["Description"] ?? "").ToString();
                        if (!string.IsNullOrEmpty(adapterNameHint) &&
                            desc.IndexOf(adapterNameHint, StringComparison.OrdinalIgnoreCase) < 0)
                            continue;

                        var ips = conf["IPAddress"] as string[];
                        var masks = conf["IPSubnet"] as string[];
                        var gws = conf["DefaultIPGateway"] as string[];
                        var dns = conf["DNSServerSearchOrder"] as string[];

                        string ip = ips?.FirstOrDefault(a => a.Contains(".")) ?? "";
                        string mask = masks?.FirstOrDefault(a => a.Contains(".")) ?? "";
                        string gw = gws?.FirstOrDefault(a => a.Contains(".")) ?? "";
                        string[] dnsArr = dns ?? new string[0];
                        return (ip, mask, gw, dnsArr);
                    }
                }
            }
            catch { }
            return ("", "", "", new string[0]);
        }
    }
}
