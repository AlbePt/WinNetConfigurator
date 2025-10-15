using System;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;

namespace WinNetConfigurator.Services
{
    /// <summary>
    /// Проверка занятости IP: ARP -> ICMP -> TCP (кастомные/дефолтные порты).
    /// Любая положительная проверка = IP занят.
    /// </summary>
    public class NetworkProbeService
    {
        // ---- ARP ----
        [DllImport("iphlpapi.dll", ExactSpelling = true)]
        private static extern int SendARP(int destIp, int srcIp, byte[] macAddr, ref int phyAddrLen);

        public bool IsAliveByArp(string ip)
        {
            try
            {
                var addr = IPAddress.Parse(ip);
                if (addr.AddressFamily != AddressFamily.InterNetwork) return false;

                int dest = BitConverter.ToInt32(addr.GetAddressBytes(), 0);
                int src = 0;
                var mac = new byte[6];
                int len = mac.Length;

                int ret = SendARP(dest, src, mac, ref len);
                if (ret != 0) return false; // 0 = NO_ERROR
                return mac.Any(b => b != 0);
            }
            catch { return false; }
        }

        // ---- ICMP ----
        public bool IsAliveByPing(string ip, int timeoutMs = 250)
        {
            try
            {
                using (var p = new Ping())
                {
                    var reply = p.Send(ip, timeoutMs, new byte[] { 0x01 }, new PingOptions(64, true));
                    return reply != null && reply.Status == IPStatus.Success;
                }
            }
            catch { return false; }
        }

        // ---- TCP ----
        private static readonly int[] _defaultPorts = new[] { 135, 139, 445, 3389 };

        public bool IsAliveByTcp(string ip, int[] customPorts = null, int timeoutMs = 150)
        {
            var ports = (customPorts == null || customPorts.Length == 0)
                ? _defaultPorts
                : customPorts.Distinct().Where(p => p > 0 && p <= 65535).ToArray();

            foreach (var port in ports)
                if (IsPortOpen(ip, port, timeoutMs)) return true;

            return false;
        }

        private bool IsPortOpen(string ip, int port, int timeoutMs)
        {
            try
            {
                using (var c = new TcpClient())
                {
                    var ar = c.BeginConnect(ip, port, null, null);
                    bool ok = ar.AsyncWaitHandle.WaitOne(TimeSpan.FromMilliseconds(timeoutMs));
                    if (!ok) return false;
                    c.EndConnect(ar);
                    return true;
                }
            }
            catch { return false; }
        }

        public bool IsIpInUse(string ip, int[] customPorts = null)
        {
            if (IsAliveByArp(ip)) return true;
            if (IsAliveByPing(ip)) return true;
            if (IsAliveByTcp(ip, customPorts)) return true;
            return false;
        }
    }
}
