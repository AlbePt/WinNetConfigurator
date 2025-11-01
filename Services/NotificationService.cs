using System;
using System.Collections.Generic;
using System.Linq;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class NotificationService
    {
        readonly List<Notification> _notifications = new List<Notification>();
        readonly object _lock = new object();

        public event EventHandler NotificationsChanged;

        public void Publish(Notification notification)
        {
            if (notification == null) throw new ArgumentNullException(nameof(notification));
            lock (_lock)
            {
                _notifications.Add(notification);
            }
            NotificationsChanged?.Invoke(this, EventArgs.Empty);
        }

        public IReadOnlyList<Notification> GetForUser(UserSession user)
        {
            lock (_lock)
            {
                return _notifications
                    .Where(n => n != null)
                    .Select(n => n)
                    .ToList();
            }
        }

        public IReadOnlyList<Notification> GetAll()
        {
            lock (_lock)
            {
                return _notifications
                    .Where(n => n != null)
                    .ToList();
            }
        }

        public void Acknowledge(Guid id)
        {
            lock (_lock)
            {
                var item = _notifications.FirstOrDefault(n => n.Id == id);
                if (item != null)
                {
                    item.Acknowledged = true;
                }
            }
            NotificationsChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
