using System;
using System.Collections.Generic;
using System.Linq;

namespace WinNetConfigurator.Models
{
    public class InventorySession
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Cabinet { get; set; }
        public UserSession Owner { get; set; }
        public DateTime StartedAt { get; set; } = DateTime.UtcNow;
        public List<InventoryEntry> Entries { get; } = new List<InventoryEntry>();
        public int CurrentIndex { get; set; }
        public bool Completed { get; set; }

        public InventoryEntry CurrentEntry =>
            Entries.Count == 0 || CurrentIndex < 0 || CurrentIndex >= Entries.Count ? null : Entries[CurrentIndex];

        public void Advance()
        {
            if (Entries.Count == 0)
                return;

            if (CurrentIndex < Entries.Count - 1)
                CurrentIndex++;
            else
                Completed = true;
        }

        public void Reset()
        {
            CurrentIndex = 0;
            Completed = false;
            foreach (var entry in Entries)
            {
                entry.Status = InventoryEntryStatus.Pending;
                entry.Notes = string.Empty;
            }
        }

        public int CountByStatus(InventoryEntryStatus status) => Entries.Count(e => e.Status == status);
    }
}
