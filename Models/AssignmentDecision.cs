using System.Collections.Generic;

namespace WinNetConfigurator.Models
{
    public class AssignmentDecision
    {
        public string SuggestedIp { get; set; }
        public List<IpPolicyWarning> Warnings { get; } = new List<IpPolicyWarning>();
        public List<RiskFactor> RiskFactors { get; } = new List<RiskFactor>();
        public bool RequiresSeniorApproval { get; set; }
        public bool HasConflicts { get; set; }
        public bool IsCabinetLimitReached { get; set; }
        public bool IsInReservation { get; set; }
    }
}
