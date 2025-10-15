using System;
using System.Net;
using System.Text.RegularExpressions;

namespace WinNetConfigurator.Utils
{
    public static class Validation
    {
        public static bool IsValidIPv4(string ip)
        {
            return IPAddress.TryParse(ip, out var addr) && addr.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork;
        }

        public static bool IsValidMask(string mask) => IsValidIPv4(mask);

        public static bool IsInRange(IPAddress ip, IPAddress start, IPAddress end)
        {
            var a = BitConverter.ToUInt32(ip.GetAddressBytes(), 0);
            var s = BitConverter.ToUInt32(start.GetAddressBytes(), 0);
            var e = BitConverter.ToUInt32(end.GetAddressBytes(), 0);
            if (BitConverter.IsLittleEndian)
            {
                a = ReverseBytes(a); s = ReverseBytes(s); e = ReverseBytes(e);
            }
            return a >= s && a <= e;
        }

        static uint ReverseBytes(uint x) => 
            (x & 0x000000FFU) << 24 | (x & 0x0000FF00U) << 8 | (x & 0x00FF0000U) >> 8 | (x & 0xFF000000U) >> 24;
    }
}
