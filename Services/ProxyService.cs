using System;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace WinNetConfigurator.Services
{
    public class ProxyService
    {
        const int INTERNET_OPTION_SETTINGS_CHANGED = 39;
        const int INTERNET_OPTION_REFRESH = 37;

        [DllImport("wininet.dll", SetLastError=true)]
        public static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);

        public void SetGlobalProxy(bool enable, string hostPort, string bypass)
        {
            // Policy: apply for all users
            using (var pol = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"))
            {
                pol.SetValue("ProxySettingsPerUser", 0, RegistryValueKind.DWord);
            }

            // HKCU for current user
            using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings"))
            {
                key.SetValue("ProxyEnable", enable ? 1 : 0, RegistryValueKind.DWord);
                if (!string.IsNullOrWhiteSpace(hostPort)) key.SetValue("ProxyServer", hostPort, RegistryValueKind.String);
                key.SetValue("ProxyOverride", bypass ?? "", RegistryValueKind.String);
            }

            // Synchronize for default user profile as baseline for new profiles
            using (var def = Registry.Users.CreateSubKey(@".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Internet Settings"))
            {
                def.SetValue("ProxyEnable", enable ? 1 : 0, RegistryValueKind.DWord);
                if (!string.IsNullOrWhiteSpace(hostPort)) def.SetValue("ProxyServer", hostPort, RegistryValueKind.String);
                def.SetValue("ProxyOverride", bypass ?? "", RegistryValueKind.String);
            }

            // Notify system
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SETTINGS_CHANGED, IntPtr.Zero, 0);
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_REFRESH, IntPtr.Zero, 0);
        }
    }
}
