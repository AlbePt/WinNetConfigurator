using System;

namespace WinNetConfigurator.Models
{
    public class Device
    {
        public int Id { get; set; }
        public int CabinetId { get; set; }
        public string CabinetName { get; set; } = string.Empty;
        public DeviceType Type { get; set; } = DeviceType.Workstation;
        public string Name { get; set; } = string.Empty;
        public string IpAddress { get; set; } = string.Empty;
        public string MacAddress { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public DateTime AssignedAt { get; set; }
    }
}
