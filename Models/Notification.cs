using System;
using System.Collections.Generic;

namespace WinNetConfigurator.Models
{
    public enum NotificationSeverity
    {
        Info,
        Warning,
        Error,
        Critical
    }

    public class Notification
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Title { get; set; }
        public string Message { get; set; }
        public NotificationSeverity Severity { get; set; }
        public string Category { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public bool RequiresAttention { get; set; }
        public bool Acknowledged { get; set; }
        public string LinkedEntityId { get; set; }
        public UserRole? RouteToRole { get; set; }
        public List<string> Tags { get; } = new List<string>();
    }
}
