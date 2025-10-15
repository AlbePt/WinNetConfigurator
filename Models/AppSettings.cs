using System;

namespace WinNetConfigurator.Models
{
    public class AppSettings
    {
        public string PoolStart { get; set; } = "";
        public string PoolEnd { get; set; } = "";
        public string Netmask { get; set; } = "";
        public string Gateway { get; set; } = "";
        public string Dns1 { get; set; } = "";
        public string Dns2 { get; set; } = "";
        public string ProxyHostPort { get; set; } = "";
        public string ProxyBypass { get; set; } = "";
        public bool ProxyGlobalOn { get; set; }
    }
}
