using System;
using System.Collections.Generic;
using System.Linq;

namespace WinNetConfigurator.Models
{
    public class AssignmentDraft
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Cabinet { get; set; }
        public string Reason { get; set; }
        public List<string> Attributes { get; } = new List<string>();
        public string RequestedIp { get; set; }
        public string SuggestedIp { get; set; }
        public AssignmentStatus Status { get; set; } = AssignmentStatus.Draft;
        public string OwnerId { get; set; }
        public string OwnerName { get; set; }
        public UserRole OwnerRole { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
        public bool RequiresSeniorApproval { get; set; }
        public List<IpPolicyWarning> Warnings { get; } = new List<IpPolicyWarning>();
        public List<RiskFactor> RiskFactors { get; } = new List<RiskFactor>();

        public AssignmentDraft Clone()
        {
            var clone = (AssignmentDraft)MemberwiseClone();
            clone.Attributes.Clear();
            clone.Attributes.AddRange(Attributes);
            clone.Warnings.Clear();
            clone.Warnings.AddRange(Warnings);
            clone.RiskFactors.Clear();
            clone.RiskFactors.AddRange(RiskFactors);
            return clone;
        }

        public void Apply(AssignmentDraft source)
        {
            Cabinet = source.Cabinet;
            Reason = source.Reason;
            Attributes.Clear();
            Attributes.AddRange(source.Attributes);
            RequestedIp = source.RequestedIp;
            SuggestedIp = source.SuggestedIp;
            Status = source.Status;
            OwnerId = source.OwnerId;
            OwnerName = source.OwnerName;
            OwnerRole = source.OwnerRole;
            RequiresSeniorApproval = source.RequiresSeniorApproval;
            Warnings.Clear();
            Warnings.AddRange(source.Warnings);
            RiskFactors.Clear();
            RiskFactors.AddRange(source.RiskFactors);
            UpdatedAt = DateTime.UtcNow;
        }

        public void SetOwner(UserSession session)
        {
            OwnerId = session?.UserId;
            OwnerName = session?.DisplayName;
            OwnerRole = session?.Role ?? UserRole.Operator;
        }

        public void SetAttributesFromText(string text)
        {
            Attributes.Clear();
            if (string.IsNullOrWhiteSpace(text))
                return;

            foreach (var token in text.Split(new[] { '\n', ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var attr = token.Trim();
                if (!string.IsNullOrEmpty(attr))
                    Attributes.Add(attr);
            }
        }

        public string ToAttributesText() => string.Join(", ", Attributes);

        public bool HasAttribute(string attr) => Attributes.Any(a => string.Equals(a, attr, StringComparison.OrdinalIgnoreCase));
    }
}
