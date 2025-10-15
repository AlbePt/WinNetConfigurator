using System;

namespace WinNetConfigurator.Models
{
    public enum InventoryEntryStatus
    {
        Pending,
        Checked,
        Missing,
        NeedsReview
    }

    public class InventoryEntry
    {
        public string DeviceId { get; set; }
        public string DisplayName { get; set; }
        public InventoryEntryStatus Status { get; set; } = InventoryEntryStatus.Pending;
        public string Notes { get; set; }
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;

        public void UpdateStatus(InventoryEntryStatus status, string notes)
        {
            Status = status;
            Notes = notes;
            UpdatedAt = DateTime.UtcNow;
        }

        public override string ToString() => $"{DisplayName} — {Status}";
    }
}
