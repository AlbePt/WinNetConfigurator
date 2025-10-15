using System;
using System.Net.NetworkInformation;

namespace WinNetConfigurator.Services
{
    public class PingService
    {
        public bool IsAlive(string ip, int timeoutMs = 300)
        {
            try
            {
                using (var p = new Ping())
                {
                    var reply = p.Send(ip, timeoutMs);
                    return reply != null && reply.Status == IPStatus.Success;
                }
            }
            catch { return false; }
        }
    }
}
