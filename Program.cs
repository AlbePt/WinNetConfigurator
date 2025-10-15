using System;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            System.IO.Directory.SetCurrentDirectory(baseDir);

            Application.ThreadException += (s, e) =>
            {
                Logger.Log("UI ThreadException: " + e.Exception);
                MessageBox.Show(e.Exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                Logger.Log("UnhandledException: " + e.ExceptionObject);
            };

            var db = new DbService();
            var network = new NetworkService();
            var exporter = new ExcelExportService();

            var settings = db.LoadSettings();
            if (!settings.IsComplete())
            {
                using (var settingsForm = new SettingsForm(settings))
                {
                    if (settingsForm.ShowDialog() != DialogResult.OK)
                        return;
                    settings = settingsForm.Result;
                    db.SaveSettings(settings);
                }
            }

            using (var cabinetForm = new CabinetSelectionForm(db.GetCabinets()))
            {
                if (cabinetForm.ShowDialog() != DialogResult.OK || cabinetForm.SelectedCabinet == null)
                    return;

                var selectedCabinet = cabinetForm.SelectedCabinet;
                if (selectedCabinet.Id == 0)
                {
                    selectedCabinet.Id = db.EnsureCabinet(selectedCabinet.Name);
                }

                RunInitialWorkflow(db, network, settings, selectedCabinet);
            }

            Application.Run(new MainForm(db, exporter));
        }

        static void RunInitialWorkflow(DbService db, NetworkService network, AppSettings settings, Cabinet cabinet)
        {
            var config = network.GetActiveConfiguration();
            if (config == null || string.IsNullOrWhiteSpace(config.IpAddress))
            {
                MessageBox.Show("Не удалось определить активный сетевой адаптер.", "Сеть", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var range = settings.GetPoolRange();
            var dnsList = string.IsNullOrWhiteSpace(settings.Dns2)
                ? new[] { settings.Dns1 }
                : new[] { settings.Dns1, settings.Dns2 };

            if (config.IsWireless)
            {
                var wifiAnswer = MessageBox.Show(
                    "Обнаружено подключение по Wi-Fi. Настройки сети не будут изменены. Записать текущий IP в базу?",
                    "Wi-Fi подключение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (wifiAnswer == DialogResult.Yes)
                {
                    SaveDevice(db, cabinet, config.IpAddress, config);
                    MessageBox.Show(
                        $"IP {config.IpAddress} записан в базу. Убедитесь, что этот адрес закреплён на маршрутизаторе.",
                        "Готово",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                return;
            }

            if (Validation.IsRouterNetwork(config.IpAddress))
            {
                OfferFreeIp(db, network, settings, cabinet, config, dnsList);
                return;
            }

            if (!Validation.IsInRange(config.IpAddress, range.Start, range.End))
            {
                var answer = MessageBox.Show(
                    "Текущий IP не входит в настроенный пул. Предложить свободный IP и применить настройки?",
                    "IP вне пула",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (answer == DialogResult.Yes)
                {
                    OfferFreeIp(db, network, settings, cabinet, config, dnsList);
                }
                return;
            }

            bool maskMatches = string.Equals(config.SubnetMask, settings.Netmask, StringComparison.OrdinalIgnoreCase);
            bool gatewayMatches = string.Equals(config.Gateway, settings.Gateway, StringComparison.OrdinalIgnoreCase);
            var currentDns = (config.Dns ?? Array.Empty<string>()).Where(Validation.IsValidIPv4).ToArray();
            bool dnsMatches = dnsList.SequenceEqual(currentDns);

            if (maskMatches && gatewayMatches && dnsMatches)
            {
                var answer = MessageBox.Show(
                    $"IP {config.IpAddress} уже настроен. Записать информацию в базу?",
                    "Запись в БД",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (answer == DialogResult.Yes)
                {
                    SaveDevice(db, cabinet, config.IpAddress, config);
                }
                return;
            }

            var fixAnswer = MessageBox.Show(
                "Некоторые сетевые параметры не совпадают с настройками. Исправить и записать в базу?",
                "Обновление настроек",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (fixAnswer == DialogResult.Yes)
            {
                network.ApplyConfiguration(config.AdapterId, config.IpAddress, settings.Netmask, settings.Gateway, dnsList);
                SaveDevice(db, cabinet, config.IpAddress, config);
            }
        }

        static void OfferFreeIp(DbService db, NetworkService network, AppSettings settings, Cabinet cabinet, NetworkConfiguration config, string[] dnsList)
        {
            var used = db.GetUsedIps();
            var available = Utils.IPRange.Enumerate(settings.PoolStart, settings.PoolEnd)
                .Select(ip => ip.ToString())
                .Where(ip => !used.Contains(ip))
                .ToList();

            if (available.Count == 0)
            {
                MessageBox.Show("Свободных IP-адресов в пуле не осталось.", "Нет адресов", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string recommended = available.First();
            using (var dialog = new FreeIpOfferForm(available, recommended))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                var ip = dialog.SelectedIp;
                network.ApplyConfiguration(config.AdapterId, ip, settings.Netmask, settings.Gateway, dnsList);
                SaveDevice(db, cabinet, ip, config);
                MessageBox.Show($"IP {ip} закреплён и применён.", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        static void SaveDevice(DbService db, Cabinet cabinet, string ip, NetworkConfiguration config)
        {
            var existing = db.GetDeviceByIp(ip);
            var device = existing ?? new Device();
            device.CabinetId = cabinet.Id;
            device.CabinetName = cabinet.Name;
            device.Type = existing?.Type ?? DeviceType.Workstation;
            device.Name = string.IsNullOrWhiteSpace(existing?.Name) ? Environment.MachineName : existing.Name;
            device.IpAddress = ip;
            device.MacAddress = string.IsNullOrWhiteSpace(config.MacAddress) ? existing?.MacAddress ?? string.Empty : config.MacAddress;
            device.Description = string.IsNullOrWhiteSpace(existing?.Description) ? config.AdapterName : existing.Description;
            device.AssignedAt = DateTime.Now;

            if (existing == null)
                db.InsertDevice(device);
            else
                db.UpdateDevice(device);
        }
    }
}
