using System;
using System.Net;
using System.Runtime.InteropServices;

namespace WinNetConfigurator.Services
{
    public class ArpService
    {
        [DllImport("iphlpapi.dll", ExactSpelling=true)]
        private static extern int SendARP(int destIp, int srcIp, byte[] macAddr, ref int phyAddrLen);

        public bool HasArpEntry(string ip)
        {
            try
            {
                var addr = IPAddress.Parse(ip);
                var bytes = addr.GetAddressBytes();
                // destIp is expected in network byte order
                int dest = BitConverter.ToInt32(bytes, 0);
                byte[] mac = new byte[6];
                int len = mac.Length;
                int res = SendARP(dest, 0, mac, ref len);
                if (res != 0) return false;
                bool allZero = true;
                for (int i = 0; i < len; i++) if (mac[i] != 0) { allZero = false; break; }
                return !allZero;
            }
            catch { return false; }
        }
    }
}
