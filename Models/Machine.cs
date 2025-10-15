using System;

namespace WinNetConfigurator.Models
{
    public class Machine
    {
        public int Id { get; set; }
        public int CabinetId { get; set; }
        public string CabinetName { get; set; } = "";
        public string Hostname { get; set; } = "";
        public string Mac { get; set; } = "";
        public string AdapterName { get; set; } = "";
        public string Ip { get; set; } = "";
        public bool ProxyOn { get; set; }
        public DateTime AssignedAt { get; set; }
        public string Source { get; set; } = "auto";
    }
}
