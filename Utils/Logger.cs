using System;
using System.IO;

namespace WinNetConfigurator.Utils
{
    public static class Logger
    {
        static readonly object _lock = new object();
        static readonly string _logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.log");

        public static void Log(string message)
        {
            lock (_lock)
            {
                File.AppendAllText(_logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\r\n");
            }
        }
    }
}
