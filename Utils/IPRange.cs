using System;
using System.Collections.Generic;
using System.Net;

namespace WinNetConfigurator.Utils
{
    public static class IPRange
    {
        public static IEnumerable<IPAddress> Enumerate(string start, string end)
        {
            var s = IPAddress.Parse(start).GetAddressBytes();
            var e = IPAddress.Parse(end).GetAddressBytes();
            if (BitConverter.IsLittleEndian) { Array.Reverse(s); Array.Reverse(e); }
            uint sInt = BitConverter.ToUInt32(s, 0);
            uint eInt = BitConverter.ToUInt32(e, 0);
            for (uint i = sInt; i <= eInt; i++)
            {
                var b = BitConverter.GetBytes(i);
                if (BitConverter.IsLittleEndian) Array.Reverse(b);
                yield return new IPAddress(b);
            }
        }
    }
}
