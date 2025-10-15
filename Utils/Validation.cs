using System;
using System.Net;
using System.Net.Sockets;

namespace WinNetConfigurator.Utils
{
    public static class Validation
    {
        public static bool IsValidIPv4(string value)
        {
            return IPAddress.TryParse(value, out var address) && address.AddressFamily == AddressFamily.InterNetwork;
        }

        public static bool IsValidMask(string value) => IsValidIPv4(value);

        public static bool IsInRange(string value, IPAddress start, IPAddress end)
        {
            if (!IPAddress.TryParse(value, out var ip)) return false;
            return IsInRange(ip, start, end);
        }

        public static bool IsInRange(IPAddress ip, IPAddress start, IPAddress end)
        {
            var ipValue = ToUInt32(ip);
            var startValue = ToUInt32(start);
            var endValue = ToUInt32(end);
            return ipValue >= startValue && ipValue <= endValue;
        }

        public static bool IsRouterNetwork(string ip)
        {
            return ip != null && ip.StartsWith("192.168.");
        }

        static uint ToUInt32(IPAddress address)
        {
            var bytes = address.GetAddressBytes();
            if (bytes.Length != 4) throw new ArgumentException("Требуется IPv4 адрес", nameof(address));
            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(bytes);
            }
            return BitConverter.ToUInt32(bytes, 0);
        }
    }
}
