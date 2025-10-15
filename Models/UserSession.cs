using System;

namespace WinNetConfigurator.Models
{
    public class UserSession
    {
        public string UserId { get; }
        public string DisplayName { get; private set; }
        public UserRole Role { get; private set; }
        public DateTime StartedAt { get; }

        public UserSession(string userId, string displayName, UserRole role)
        {
            if (string.IsNullOrWhiteSpace(userId))
                throw new ArgumentException("User identifier is required", nameof(userId));
            UserId = userId;
            DisplayName = string.IsNullOrWhiteSpace(displayName) ? userId : displayName.Trim();
            Role = role;
            StartedAt = DateTime.UtcNow;
        }

        public void UpdateDisplayName(string displayName)
        {
            if (!string.IsNullOrWhiteSpace(displayName))
            {
                DisplayName = displayName.Trim();
            }
        }

        public void UpdateRole(UserRole role)
        {
            Role = role;
        }

        public override string ToString() => $"{DisplayName} ({Role})";
    }
}
