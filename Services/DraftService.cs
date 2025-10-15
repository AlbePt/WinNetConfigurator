using System;
using System.Collections.Generic;
using System.Linq;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class DraftService
    {
        readonly List<AssignmentDraft> _drafts = new List<AssignmentDraft>();
        readonly object _lock = new object();

        public event EventHandler DraftsChanged;

        public AssignmentDraft SaveDraft(UserSession user, AssignmentDraft draft)
        {
            if (user == null || draft == null)
                throw new ArgumentNullException();

            lock (_lock)
            {
                var existing = _drafts.FirstOrDefault(d => d.Id == draft.Id);
                draft.SetOwner(user);
                draft.UpdatedAt = DateTime.UtcNow;
                if (existing == null)
                {
                    _drafts.Add(draft.Clone());
                }
                else
                {
                    existing.Apply(draft);
                }
            }
            DraftsChanged?.Invoke(this, EventArgs.Empty);
            return draft;
        }

        public IEnumerable<AssignmentDraft> GetDrafts(UserSession user, bool includeApproved = false)
        {
            lock (_lock)
            {
                return _drafts
                    .Where(d => includeApproved || d.Status == AssignmentStatus.Draft || d.Status == AssignmentStatus.PendingApproval)
                    .Where(d => user == null || string.Equals(d.OwnerId, user.UserId, StringComparison.OrdinalIgnoreCase))
                    .Select(d => d.Clone())
                    .ToList();
            }
        }

        public void UpdateStatus(Guid draftId, AssignmentStatus status)
        {
            lock (_lock)
            {
                var existing = _drafts.FirstOrDefault(d => d.Id == draftId);
                if (existing != null)
                {
                    existing.Status = status;
                    existing.UpdatedAt = DateTime.UtcNow;
                }
            }
            DraftsChanged?.Invoke(this, EventArgs.Empty);
        }

        public AssignmentDraft GetDraft(Guid id)
        {
            lock (_lock)
            {
                var existing = _drafts.FirstOrDefault(d => d.Id == id);
                return existing?.Clone();
            }
        }
    }
}
