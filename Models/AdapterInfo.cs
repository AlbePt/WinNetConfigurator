using System;

namespace WinNetConfigurator.Models
{
    public class AdapterInfo
    {
        public string NetConnectionId { get; set; } = "";
        public string Name { get; set; } = "";
        public string MacAddress { get; set; } = "";
        public string IPv4Address { get; set; } = "";
        public string Type { get; set; } = ""; // Ethernet/Wi-Fi/Other
        public bool IsUp { get; set; }
    }
}
