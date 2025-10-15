using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace WinNetConfigurator.Models
{
    public class IpPolicyState
    {
        public int DefaultCabinetLimit { get; set; } = 10;
        public Dictionary<string, CabinetPolicy> CabinetPolicies { get; } = new Dictionary<string, CabinetPolicy>(StringComparer.OrdinalIgnoreCase);
        public List<IpReservation> Reservations { get; } = new List<IpReservation>();
        public List<PresumedUsage> PresumedAddresses { get; } = new List<PresumedUsage>();

        public CabinetPolicy GetPolicyForCabinet(string cabinet)
        {
            if (string.IsNullOrWhiteSpace(cabinet))
            {
                return null;
            }
            CabinetPolicies.TryGetValue(cabinet.Trim(), out var policy);
            return policy;
        }
    }

    public class CabinetPolicy
    {
        public string CabinetName { get; set; }
        public int Limit { get; set; }
        public Dictionary<string, int> SubPoolLimits { get; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> Exceptions { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    public class IpReservation
    {
        public string RangeStart { get; set; }
        public string RangeEnd { get; set; }
        public string Reason { get; set; }
        public bool HardBlock { get; set; }

        public bool Contains(string ip)
        {
            if (string.IsNullOrWhiteSpace(ip) || string.IsNullOrWhiteSpace(RangeStart))
                return false;

            var value = ToUInt32(IPAddress.Parse(ip));
            var start = ToUInt32(IPAddress.Parse(RangeStart));
            var end = ToUInt32(IPAddress.Parse(string.IsNullOrWhiteSpace(RangeEnd) ? RangeStart : RangeEnd));
            if (start > end)
            {
                var tmp = start;
                start = end;
                end = tmp;
            }
            return value >= start && value <= end;
        }

        static uint ToUInt32(IPAddress address)
        {
            var bytes = address.GetAddressBytes();
            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(bytes);
            }
            return BitConverter.ToUInt32(bytes, 0);
        }
    }

    public class PresumedUsage
    {
        public string Ip { get; set; }
        public string Source { get; set; }
        public DateTime ObservedAt { get; set; }
        public bool RequiresConfirmation { get; set; } = true;
    }

    public class IpPolicyWarning
    {
        public string Code { get; set; }
        public string Title { get; set; }
        public string Message { get; set; }
        public bool RequiresSeniorApproval { get; set; }
        public RiskFactor RiskFactor { get; set; }

        public override string ToString() => string.IsNullOrEmpty(Title) ? Message : $"{Title}: {Message}";
    }

    public enum RiskFactor
    {
        None,
        Limit,
        Conflict,
        Reservation,
        PresumedOccupied,
        MissingAttributes,
        UnknownCabinet,
        PolicyException
    }
}
