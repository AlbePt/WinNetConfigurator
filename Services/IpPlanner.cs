using System;
using System.Net;
using WinNetConfigurator.Utils;
using System.Collections.Generic;

namespace WinNetConfigurator.Services
{
    public class IpPlanner
    {
        public string SuggestNextFreeIp(string poolStart, string poolEnd, Func<string, bool> isBusy)
        {
            foreach (var ip in IPRange.Enumerate(poolStart, poolEnd))
            {
                var s = ip.ToString();
                if (isBusy(s)) continue;
                return s;
            }
            throw new InvalidOperationException("Свободных IP не найдено в заданном диапазоне.");
        }
    }
}
