using System;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class RoleAccessService
    {
        public bool CanEditDraft(UserSession user, AssignmentDraft draft)
        {
            if (user == null || draft == null)
                return false;
            if (user.Role == UserRole.Administrator)
                return true;
            if (user.Role == UserRole.Auditor)
                return false;
            if (string.Equals(draft.OwnerId, user.UserId, StringComparison.OrdinalIgnoreCase))
                return true;
            if (user.Role == UserRole.SeniorOperator && draft.Status != AssignmentStatus.Approved)
                return true;
            return false;
        }

        public bool CanApprove(UserSession user, AssignmentDraft draft)
        {
            if (user == null || draft == null)
                return false;
            return user.Role == UserRole.SeniorOperator || user.Role == UserRole.Administrator;
        }

        public bool CanView(UserSession user, AssignmentDraft draft)
        {
            if (user == null || draft == null)
                return false;
            if (user.Role == UserRole.Administrator || user.Role == UserRole.SeniorOperator)
                return true;
            if (user.Role == UserRole.Auditor)
                return draft.Status == AssignmentStatus.Approved;
            return string.Equals(draft.OwnerId, user.UserId, StringComparison.OrdinalIgnoreCase) || draft.Status == AssignmentStatus.Approved;
        }

        public bool CanSeeNotification(UserSession user, Notification notification)
        {
            if (user == null || notification == null)
                return false;
            if (notification.RouteToRole.HasValue)
                return notification.RouteToRole.Value == user.Role;
            if (notification.Severity == NotificationSeverity.Critical)
                return true;
            if (user.Role == UserRole.Administrator)
                return true;
            if (notification.RequiresAttention && user.Role == UserRole.SeniorOperator)
                return true;
            return user.Role == UserRole.Operator && notification.Severity <= NotificationSeverity.Warning;
        }
    }
}
