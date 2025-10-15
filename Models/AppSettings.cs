using System;
using System.Net;

namespace WinNetConfigurator.Models
{
    public class AppSettings
    {
        public string PoolStart { get; set; } = string.Empty;
        public string PoolEnd { get; set; } = string.Empty;
        public string Netmask { get; set; } = string.Empty;
        public string Gateway { get; set; } = string.Empty;
        public string Dns1 { get; set; } = string.Empty;
        public string Dns2 { get; set; } = string.Empty;

        public bool IsComplete()
        {
            return !string.IsNullOrWhiteSpace(PoolStart)
                && !string.IsNullOrWhiteSpace(PoolEnd)
                && !string.IsNullOrWhiteSpace(Netmask)
                && !string.IsNullOrWhiteSpace(Gateway)
                && !string.IsNullOrWhiteSpace(Dns1);
        }

        public (IPAddress Start, IPAddress End) GetPoolRange()
        {
            if (!IPAddress.TryParse(PoolStart, out var start) || !IPAddress.TryParse(PoolEnd, out var end))
                throw new InvalidOperationException("Диапазон IP настроен некорректно.");
            return (start, end);
        }
    }
}
