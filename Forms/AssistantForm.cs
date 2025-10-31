using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator.Forms
{
    public class AssistantForm : Form
    {
        readonly DbService db;
        readonly NetworkService network;
        readonly AppSettings settings;

        readonly ComboBox cmbCabinets = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        readonly Button btnReloadCabinets = new Button { Text = "Обновить", AutoSize = true };
        readonly Button btnNewCabinet = new Button { Text = "Новый кабинет…", AutoSize = true };
        readonly Button btnRun = new Button { Text = "Считать и записать этот ПК", AutoSize = true };
        readonly TextBox txtLog = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };

        public AssistantForm(DbService db, NetworkService network, AppSettings settings)
        {
            this.db = db ?? throw new ArgumentNullException(nameof(db));
            this.network = network ?? throw new ArgumentNullException(nameof(network));
            this.settings = settings ?? new AppSettings();

            Text = "WinNetConfigurator — режим ассистента";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 700;
            Height = 500;

            BuildLayout();
            LoadCabinets();
        }

        void BuildLayout()
        {
            var top = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 70,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(12, 10, 12, 6)
            };
            top.Controls.Add(new Label { Text = "Кабинет:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
            top.Controls.Add(cmbCabinets);
            top.Controls.Add(btnReloadCabinets);
            top.Controls.Add(btnNewCabinet);

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 56,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(12, 2, 12, 6)
            };
            btnRun.Font = new Font(Font.FontFamily, 10, FontStyle.Bold);
            btnRun.Height = 36;
            actions.Controls.Add(btnRun);

            Controls.Add(txtLog);
            Controls.Add(actions);
            Controls.Add(top);

            btnReloadCabinets.Click += (_, __) => LoadCabinets();
            btnNewCabinet.Click += (_, __) => CreateCabinet();
            btnRun.Click += (_, __) => RunForSelectedCabinet();
        }

        void Log(string message)
        {
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
        }

        void LoadCabinets()
        {
            var list = db.GetCabinets();
            cmbCabinets.Items.Clear();
            foreach (var c in list)
                cmbCabinets.Items.Add(c);
            cmbCabinets.DisplayMember = "Name";
            if (cmbCabinets.Items.Count > 0)
                cmbCabinets.SelectedIndex = 0;
            Log("Загружены кабинеты из БД.");
        }

        void CreateCabinet()
        {
            var existing = db.GetCabinets();
            using (var dialog = new CabinetSelectionForm(existing))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                var name = dialog.CabinetName;
                if (string.IsNullOrWhiteSpace(name))
                    return;

                try
                {
                    var id = db.EnsureCabinet(name);
                    Log($"Кабинет «{name}» сохранён (id={id}).");
                    LoadCabinets();
                    for (int i = 0; i < cmbCabinets.Items.Count; i++)
                    {
                        if (((Cabinet)cmbCabinets.Items[i]).Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                        {
                            cmbCabinets.SelectedIndex = i;
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Не удалось сохранить кабинет: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Log("Ошибка сохранения кабинета: " + ex.Message);
                }
            }
        }

        void RunForSelectedCabinet()
        {
            if (cmbCabinets.SelectedItem is not Cabinet cab)
            {
                MessageBox.Show("Выберите кабинет.", "Нет кабинета", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (RunInitialWorkflow(cab))
            {
                Log($"ПК в кабинете «{cab.Name}» записан в БД.");
            }
        }

        bool RunInitialWorkflow(Cabinet cabinet)
        {
            var config = network.GetActiveConfiguration();
            if (config == null || string.IsNullOrWhiteSpace(config.IpAddress))
            {
                MessageBox.Show(
                    "Не удалось определить активный сетевой адаптер.",
                    "Сеть",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                Log("Не удалось определить активный сетевой адаптер.");
                return false;
            }

            var dnsList = string.IsNullOrWhiteSpace(settings.Dns2)
                ? new[] { settings.Dns1 }
                : new[] { settings.Dns1, settings.Dns2 };

            if (Validation.IsRouterNetwork(config.IpAddress))
            {
                return HandleRouterNetwork(cabinet, config, dnsList);
            }

            (System.Net.IPAddress Start, System.Net.IPAddress End) range;
            try
            {
                range = settings.GetPoolRange();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Диапазон IP настроен некорректно: " + ex.Message,
                    "Настройки сети",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Log("Диапазон IP некорректен: " + ex.Message);
                return false;
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
                    if (!TrySelectFreeIp("Выберите IP из пула для этого ПК", out var selectedIp))
                        return false;

                    try
                    {
                        network.ApplyConfiguration(config.AdapterId, selectedIp, settings.Netmask, settings.Gateway, dnsList);
                        SaveDevice(cabinet, selectedIp, config);
                        MessageBox.Show("Настройки применены и записаны.", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Log($"Настройки применены и записаны. IP={selectedIp}");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Не удалось применить настройки: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Log("Ошибка применения настроек: " + ex.Message);
                        return false;
                    }
                }

                return false;
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
                    SaveDevice(cabinet, config.IpAddress, config);
                    Log($"IP {config.IpAddress} записан в базу без изменения настроек.");
                    return true;
                }
                return false;
            }

            var fixAnswer = MessageBox.Show(
                "Некоторые сетевые параметры не совпадают с настройками. Исправить и записать в базу?",
                "Обновление настроек",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (fixAnswer == DialogResult.Yes)
            {
                try
                {
                    network.ApplyConfiguration(config.AdapterId, config.IpAddress, settings.Netmask, settings.Gateway, dnsList);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Не удалось применить настройки сети: " + ex.Message,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    Log("Не удалось применить настройки сети: " + ex.Message);
                    return false;
                }

                SaveDevice(cabinet, config.IpAddress, config);
                MessageBox.Show("Настройки обновлены и записаны.", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Log("Настройки обновлены и записаны.");
                return true;
            }

            return false;
        }

        bool HandleRouterNetwork(Cabinet cabinet, NetworkConfiguration config, string[] dnsList)
        {
            using (var dialog = new RouterIpActionForm(config.IpAddress))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return false;

                switch (dialog.SelectedAction)
                {
                    case RouterIpAction.BindOnly:
                        if (!TrySelectFreeIp("Выберите свободный IP из пула для закрепления в базе.", out var selectedIp))
                            return false;

                        SaveDevice(cabinet, selectedIp, config);
                        MessageBox.Show($"IP {selectedIp} закреплён в базе. Настройки компьютера не изменялись.", "Готово",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Log($"IP {selectedIp} закреплён в базе (роутер-сеть).");
                        return true;

                    case RouterIpAction.BindAndConfigure:
                        if (!TrySelectFreeIp("Выберите свободный IP для назначения этому ПК", out var selectedIp2))
                            return false;

                        try
                        {
                            network.ApplyConfiguration(config.AdapterId, selectedIp2, settings.Netmask, settings.Gateway, dnsList);
                            SaveDevice(cabinet, selectedIp2, config);
                            MessageBox.Show($"IP {selectedIp2} закреплён и применён.", "Готово",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            Log($"IP {selectedIp2} закреплён и применён (роутер-сеть).");
                            return true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Не удалось применить настройки сети: " + ex.Message, "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Log("Ошибка применения сетевых настроек: " + ex.Message);
                            return false;
                        }

                    default:
                        return false;
                }
            }
        }

        bool TrySelectFreeIp(string description, out string selectedIp)
        {
            selectedIp = null;
            var used = db.GetUsedIps();
            using (var dialog = new FreeIpOfferForm(
                description,
                settings.PoolStart,
                settings.PoolEnd,
                used,
                allowCustomEntry: true))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return false;
                selectedIp = dialog.SelectedIp;
                return true;
            }
        }

        void SaveDevice(Cabinet cabinet, string ip, NetworkConfiguration config)
        {
            var existing = db.GetDeviceByIp(ip);
            var device = existing ?? new Device();
            device.CabinetId = cabinet.Id;
            device.CabinetName = cabinet.Name;
            device.Type = existing?.Type ?? DeviceType.Workstation;
            device.Name = string.IsNullOrWhiteSpace(existing?.Name) ? Environment.MachineName : existing.Name;
            device.IpAddress = ip;
            device.MacAddress = string.IsNullOrWhiteSpace(config.MacAddress)
                ? existing?.MacAddress ?? string.Empty
                : config.MacAddress;
            device.Description = string.IsNullOrWhiteSpace(existing?.Description)
                ? config.AdapterName
                : existing.Description;
            device.AssignedAt = DateTime.Now;

            if (existing == null)
                db.InsertDevice(device);
            else
                db.UpdateDevice(device);
        }
    }
}
