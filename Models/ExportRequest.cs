using System;
using System.Collections.Generic;

namespace WinNetConfigurator.Models
{
    public class ExportRequest
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public List<string> Locations { get; } = new List<string>();
        public DateTime? From { get; set; }
        public DateTime? To { get; set; }
        public List<string> Statuses { get; } = new List<string>();
        public string Responsible { get; set; }
        public string Format { get; set; } = "xlsx";
        public bool IncludeDrafts { get; set; }
        public UserSession RequestedBy { get; set; }
        public DateTime RequestedAt { get; set; } = DateTime.UtcNow;
    }
}
