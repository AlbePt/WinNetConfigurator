namespace WinNetConfigurator.Models
{
    public class NetworkConfiguration
    {
        public string AdapterId { get; set; } = string.Empty;
        public string AdapterName { get; set; } = string.Empty;
        public string MacAddress { get; set; } = string.Empty;
        public string IpAddress { get; set; } = string.Empty;
        public string SubnetMask { get; set; } = string.Empty;
        public string Gateway { get; set; } = string.Empty;
        public string[] Dns { get; set; } = new string[0];
        public bool IsDhcpEnabled { get; set; }
        public bool IsWireless { get; set; }
    }
}
