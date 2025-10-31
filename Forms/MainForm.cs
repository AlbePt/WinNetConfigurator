using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator.Forms
{
    public class MainForm : Form
    {
        readonly DbService db;
        readonly ExcelExportService excel;
        readonly NetworkService network;
        readonly BindingList<Device> devices = new BindingList<Device>();
        readonly DataGridView grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false };
        readonly ContextMenuStrip settingsMenu = new ContextMenuStrip();
        readonly CabinetNameComparer cabinetComparer = new CabinetNameComparer();
        readonly IpAddressComparer ipAddressComparer = new IpAddressComparer();
        Button btnSettings;
        string currentSortProperty;
        bool sortAscending = true;
        AppSettings settings;

        readonly Panel assistantPanel = new Panel();
        readonly Label assistantTitle = new Label();
        readonly Label assistantText = new Label();

        const string SettingsPassword = "3baRTfg6";

        public MainForm(DbService db, ExcelExportService excelService, NetworkService networkService, AppSettings initialSettings)
        {
            this.db = db;
            excel = excelService;
            network = networkService;
            settings = initialSettings ?? new AppSettings();

            currentSortProperty = nameof(Device.CabinetName);
            sortAscending = true;

            Text = "WinNetConfigurator";
            Width = 960;
            Height = 600;
            StartPosition = FormStartPosition.CenterScreen;

            BuildLayout();
            LoadDevices();
        }

        void BuildLayout()
        {
            var panelButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 48,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(12, 10, 12, 6)
            };

            var btnAdd = new Button { Text = "Добавить устройство", AutoSize = true };
            var btnEdit = new Button { Text = "Редактировать", AutoSize = true };
            var btnDelete = new Button { Text = "Удалить", AutoSize = true };
            var btnRefresh = new Button { Text = "Обновить", AutoSize = true };
            var btnExport = new Button { Text = "Экспорт в XLSX", AutoSize = true };
            var btnImportXlsx = new Button { Text = "Импорт из XLSX", AutoSize = true };
            btnSettings = new Button { Text = "Настройки", AutoSize = true };

            btnAdd.Click += (_, __) => AddDevice();
            btnEdit.Click += (_, __) => EditSelected();
            btnDelete.Click += (_, __) => DeleteSelected();
            btnRefresh.Click += (_, __) => LoadDevices();
            btnExport.Click += (_, __) => Export();
            btnImportXlsx.Click += (_, __) => ImportFromExcel();
            btnSettings.Click += (_, __) => ShowSettingsMenu();

            InitializeSettingsMenu();

            panelButtons.Controls.AddRange(new Control[]
            {
                btnAdd,
                btnEdit,
                btnDelete,
                btnRefresh,
                btnExport,
                btnImportXlsx,
                btnSettings
            });

            assistantPanel.Dock = DockStyle.Top;
            assistantPanel.Height = 56;
            assistantPanel.Padding = new Padding(12, 6, 12, 6);
            assistantPanel.BackColor = Color.FromArgb(255, 252, 230);

            assistantTitle.AutoSize = true;
            assistantTitle.Font = new Font(Font, FontStyle.Bold);
            assistantTitle.Text = "Подсказка";
            assistantTitle.Location = new Point(0, 0);

            assistantText.AutoSize = true;
            assistantText.Width = 600;
            assistantText.MaximumSize = new Size(800, 0);

            assistantPanel.Controls.Add(assistantTitle);
            assistantPanel.Controls.Add(assistantText);
            assistantText.Left = 0;
            assistantText.Top = assistantTitle.Bottom + 2;

            grid.DataSource = devices;

            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "№",
                Name = "RowNumber",
                Width = 50,
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleRight
                }
            });
            grid.Columns.Add(CreateColumn("Тип", nameof(Device.Type), 120));
            grid.Columns.Add(CreateColumn("Кабинет", nameof(Device.CabinetName), 140));
            grid.Columns.Add(CreateColumn("Название", nameof(Device.Name), 180));
            grid.Columns.Add(CreateColumn("IP адрес", nameof(Device.IpAddress), 120));
            grid.Columns.Add(CreateColumn("MAC адрес", nameof(Device.MacAddress), 140));
            grid.Columns.Add(CreateColumn("Описание", nameof(Device.Description), -1));
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Закреплено",
                DataPropertyName = nameof(Device.AssignedAt),
                Width = 150,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "dd.MM.yyyy HH:mm" },
                SortMode = DataGridViewColumnSortMode.Programmatic
            });
            grid.CellFormatting += Grid_CellFormatting;
            grid.ColumnHeaderMouseClick += Grid_ColumnHeaderMouseClick;

            Controls.Add(grid);
            Controls.Add(assistantPanel);
            Controls.Add(panelButtons);
        }

        DataGridViewColumn CreateColumn(string header, string property, int width)
        {
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = header,
                DataPropertyName = property,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };

            if (width > 0)
                column.Width = width;
            else
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            return column;
        }

        void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (grid.Columns[e.ColumnIndex].Name == "RowNumber")
            {
                e.Value = (e.RowIndex + 1).ToString();
                e.FormattingApplied = true;
                return;
            }

            if (grid.Columns[e.ColumnIndex].DataPropertyName == nameof(Device.Type) && e.Value is DeviceType type)
            {
                switch (type)
                {
                    case DeviceType.Printer:
                        e.Value = "Принтер";
                        break;
                    case DeviceType.Other:
                        e.Value = "Другое устройство";
                        break;
                    default:
                        e.Value = "Рабочее место";
                        break;
                }
                e.FormattingApplied = true;
            }
        }

        void ShowAssistant(string scenario)
        {
            switch (scenario)
            {
                case "empty":
                    assistantTitle.Text = "Что сделать первым?";
                    assistantText.Text =
                        "Сейчас в базе нет устройств. Нажмите «Добавить устройство» или «Импорт из XLSX», " +
                        "чтобы заполнить список.";
                    break;

                case "devices_loaded":
                    assistantTitle.Text = "Подсказка";
                    assistantText.Text =
                        "Двойной щелчок по строке — редактирование. Правой кнопкой — действия. " +
                        "Для обновления нажмите «Обновить».";
                    break;

                case "after_import":
                    assistantTitle.Text = "Импорт завершён";
                    assistantText.Text =
                        "Проверьте, что кабинеты и IP распределены корректно. При необходимости откройте «Настройки».";
                    break;

                default:
                    assistantTitle.Text = "Подсказка";
                    assistantText.Text = "Выберите строку, чтобы редактировать устройство или удалить его.";
                    break;
            }
        }

        void LoadDevices()
        {
            devices.Clear();
            foreach (var device in db.GetDevices())
                devices.Add(device);
            ApplySort(currentSortProperty, preserveDirection: true);
            if (!string.IsNullOrEmpty(currentSortProperty))
            {
                var column = grid.Columns.Cast<DataGridViewColumn>()
                    .FirstOrDefault(c => c.DataPropertyName == currentSortProperty);
                if (column != null)
                    UpdateSortGlyph(column);
            }

            ShowAssistant(devices.Count == 0 ? "empty" : "devices_loaded");
        }

        void InitializeSettingsMenu()
        {
            settingsMenu.Items.Clear();
            settingsMenu.Items.Add("Настройки сети", null, (_, __) => EditSettings());
            settingsMenu.Items.Add(new ToolStripSeparator());
            settingsMenu.Items.Add("Импорт БД", null, (_, __) => ImportDatabase());
            settingsMenu.Items.Add("Резервная копия БД", null, (_, __) => BackupDatabase());
            settingsMenu.Items.Add("Очистить БД", null, (_, __) => ClearDatabase());
        }

        void ShowSettingsMenu()
        {
            if (btnSettings == null)
                return;

            if (!EnsureSettingsPassword())
                return;

            var location = new Point(0, btnSettings.Height);
            settingsMenu.Show(btnSettings, location);
        }

        bool EnsureSettingsPassword()
        {
            using (var dialog = new PasswordPromptForm())
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return false;

                if (!string.Equals(dialog.EnteredPassword, SettingsPassword, StringComparison.Ordinal))
                {
                    MessageBox.Show("Неверный пароль.", "Доступ запрещён", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            return true;
        }

        void Grid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var column = grid.Columns[e.ColumnIndex];
            if (string.IsNullOrEmpty(column.DataPropertyName))
                return;

            ApplySort(column.DataPropertyName);
            UpdateSortGlyph(column);
        }

        void ApplySort(string propertyName, bool preserveDirection = false)
        {
            if (string.IsNullOrEmpty(propertyName))
                return;

            if (!preserveDirection)
            {
                if (string.Equals(currentSortProperty, propertyName, StringComparison.Ordinal))
                    sortAscending = !sortAscending;
                else
                    sortAscending = true;
            }

            currentSortProperty = propertyName;

            IEnumerable<Device> ordered;
            switch (propertyName)
            {
                case nameof(Device.CabinetName):
                    ordered = sortAscending
                        ? devices.OrderBy(d => d.CabinetName, cabinetComparer)
                        : devices.OrderByDescending(d => d.CabinetName, cabinetComparer);
                    break;
                case nameof(Device.IpAddress):
                    ordered = sortAscending
                        ? devices.OrderBy(d => d.IpAddress, ipAddressComparer)
                        : devices.OrderByDescending(d => d.IpAddress, ipAddressComparer);
                    break;
                default:
                    var property = TypeDescriptor.GetProperties(typeof(Device))[propertyName];
                    if (property == null)
                        return;
                    ordered = sortAscending
                        ? devices.OrderBy(d => property.GetValue(d))
                        : devices.OrderByDescending(d => property.GetValue(d));
                    break;
            }

            var sortedList = ordered.ToList();

            devices.Clear();
            foreach (var device in sortedList)
                devices.Add(device);
        }

        void UpdateSortGlyph(DataGridViewColumn sortedColumn)
        {
            foreach (DataGridViewColumn column in grid.Columns)
            {
                column.HeaderCell.SortGlyphDirection = column == sortedColumn
                    ? (sortAscending ? SortOrder.Ascending : SortOrder.Descending)
                    : SortOrder.None;
            }
        }

        Device GetSelectedDevice()
        {
            if (grid.CurrentRow?.DataBoundItem is Device device)
                return device;
            return null;
        }

        void AddDevice()
        {
            settings = db.LoadSettings();
            if (!settings.IsComplete())
            {
                MessageBox.Show(
                    "Сначала заполните настройки сети в разделе \"Настройки\".",
                    "Настройки сети",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            var cabinets = db.GetCabinets();
            using (var dialog = new CabinetSelectionForm(cabinets))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                var cabinetName = dialog.CabinetName;
                if (string.IsNullOrWhiteSpace(cabinetName))
                    return;

                int cabinetId;
                try
                {
                    cabinetId = db.EnsureCabinet(cabinetName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Не удалось сохранить кабинет: " + ex.Message,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                var cabinet = new Cabinet { Id = cabinetId, Name = cabinetName };
                if (RunInitialWorkflow(cabinet))
                    LoadDevices();
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
                return false;
            }

            var dnsList = string.IsNullOrWhiteSpace(settings.Dns2)
                ? new[] { settings.Dns1 }
                : new[] { settings.Dns1, settings.Dns2 };

            if (config.IsWireless)
            {
                const string message =
                    "Обнаружено подключение по Wi-Fi. Введите IP-адрес, закреплённый на маршрутизаторе для этого рабочего места. Запись будет сохранена только в базе данных.";

                using (var dialog = new ManualIpEntryForm(message, config.IpAddress))
                {
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        SaveDevice(cabinet, dialog.EnteredIp, config);
                        MessageBox.Show(
                            $"IP {dialog.EnteredIp} записан в базу. Настройки компьютера не изменялись.",
                            "Готово",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return true;
                    }
                }
                return false;
            }

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
                    return OfferFreeIp(cabinet, config, dnsList);
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
                    return false;
                }

                SaveDevice(cabinet, config.IpAddress, config);
                return true;
            }

            return false;
        }

        bool TrySelectFreeIp(string prompt, out string selectedIp)
        {
            selectedIp = null;

            var used = new HashSet<string>(db.GetUsedIps());
            List<string> available;
            try
            {
                available = IPRange.Enumerate(settings.PoolStart, settings.PoolEnd)
                    .Select(ip => ip.ToString())
                    .Where(ip => !used.Contains(ip))
                    .ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Не удалось получить свободные адреса: " + ex.Message,
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }

            if (available.Count == 0)
            {
                MessageBox.Show(
                    "Свободных IP-адресов в пуле не осталось.",
                    "Нет адресов",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            string recommended = available.First();
            using (var dialog = new FreeIpOfferForm(available, recommended, prompt))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return false;

                selectedIp = dialog.SelectedIp;
                return true;
            }
        }

        bool OfferFreeIp(Cabinet cabinet, NetworkConfiguration config, string[] dnsList)
        {
            if (!TrySelectFreeIp(
                "Выберите свободный IP из пула и примените его к этому рабочему месту.",
                out var ip))
            {
                return false;
            }

            try
            {
                network.ApplyConfiguration(config.AdapterId, ip, settings.Netmask, settings.Gateway, dnsList);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Не удалось применить настройки сети: " + ex.Message,
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }

            SaveDevice(cabinet, ip, config);
            MessageBox.Show(
                $"IP {ip} закреплён и применён.",
                "Готово",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return true;
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
                        if (!TrySelectFreeIp(
                            "Выберите свободный IP из пула для закрепления в базе.",
                            out var selectedIp))
                        {
                            return false;
                        }

                        SaveDevice(cabinet, selectedIp, config);
                        MessageBox.Show(
                            $"IP {selectedIp} закреплён в базе. Настройки компьютера не изменялись.",
                            "Готово",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return true;
                    case RouterIpAction.BindAndConfigure:
                        return OfferFreeIp(cabinet, config, dnsList);
                    default:
                        return false;
                }
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
            device.MacAddress = string.IsNullOrWhiteSpace(config.MacAddress) ? existing?.MacAddress ?? string.Empty : config.MacAddress;
            device.Description = string.IsNullOrWhiteSpace(existing?.Description) ? config.AdapterName : existing.Description;
            device.AssignedAt = DateTime.Now;

            if (existing == null)
                db.InsertDevice(device);
            else
                db.UpdateDevice(device);
        }

        void EditSelected()
        {
            var selected = GetSelectedDevice();
            if (selected == null)
            {
                MessageBox.Show("Выберите устройство в таблице.", "Нет выбора", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var cabinets = db.GetCabinets().ToArray();
            using (var dialog = new DeviceEditForm(db, new Device
            {
                Id = selected.Id,
                CabinetId = selected.CabinetId,
                CabinetName = selected.CabinetName,
                Type = selected.Type,
                Name = selected.Name,
                IpAddress = selected.IpAddress,
                MacAddress = selected.MacAddress,
                Description = selected.Description,
                AssignedAt = selected.AssignedAt
            }, cabinets))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var conflict = db.GetDeviceByIp(dialog.Result.IpAddress);
                    if (conflict != null && conflict.Id != selected.Id)
                    {
                        MessageBox.Show("Такой IP уже зарегистрирован в базе.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    dialog.Result.Id = selected.Id;
                    dialog.Result.AssignedAt = DateTime.Now;
                    db.UpdateDevice(dialog.Result);
                    LoadDevices();
                }
            }
        }

        void DeleteSelected()
        {
            var selected = GetSelectedDevice();
            if (selected == null)
            {
                MessageBox.Show("Выберите устройство в таблице.", "Нет выбора", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var answer = MessageBox.Show($"Удалить устройство {selected.Name} ({selected.IpAddress})?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (answer != DialogResult.Yes)
                return;

            if (db.DeleteDevice(selected.Id))
            {
                devices.Remove(selected);
            }
            else
            {
                MessageBox.Show(
                    "Не удалось удалить устройство из базы данных.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        void Export()
        {
            using (var dialog = new SaveFileDialog { Filter = "Excel файлы (*.xlsx)|*.xlsx", FileName = $"export_{DateTime.Now:yyyyMMdd_HHmm}.xlsx" })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    excel.ExportDevices(devices, dialog.FileName);
                    MessageBox.Show("Экспорт завершён успешно.", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        void ImportFromExcel()
        {
            using (var dialog = new OpenFileDialog { Filter = "Excel файлы (*.xlsx)|*.xlsx" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                var confirm = MessageBox.Show(
                    "Текущие устройства и кабинеты будут заменены данными из файла. Продолжить?",
                    "Импорт из XLSX",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);

                if (confirm != DialogResult.Yes)
                    return;

                try
                {
                    var imported = excel.ImportDevices(dialog.FileName);
                    if (imported.Count == 0)
                    {
                        MessageBox.Show(
                            "В выбранном файле не найдено устройств для импорта.",
                            "Импорт из XLSX",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }

                    db.ReplaceDevices(imported);
                    LoadDevices();

                    MessageBox.Show(
                        "Импорт из XLSX выполнен успешно.",
                        "Импорт из XLSX",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Не удалось выполнить импорт из XLSX: " + ex.Message,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        void EditSettings()
        {
            var current = db.LoadSettings();
            using (var dialog = new SettingsForm(current))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    db.SaveSettings(dialog.Result);
                    settings = dialog.Result;
                    MessageBox.Show("Настройки обновлены.", "Настройки", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        void ClearDatabase()
        {
            var answer = MessageBox.Show(
                "Это удалит все устройства, кабинеты и настройки сети. Продолжить?",
                "Очистка базы данных",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            if (answer != DialogResult.Yes)
                return;

            db.ClearDatabase();
            LoadDevices();
            settings = db.LoadSettings();
            MessageBox.Show("База данных очищена.", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void ImportDatabase()
        {
            using (var dialog = new OpenFileDialog { Filter = "Файлы базы данных (*.db)|*.db|Все файлы (*.*)|*.*" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                var confirm = MessageBox.Show(
                    "Текущая база данных будет заменена выбранным файлом. Продолжить?",
                    "Импорт базы данных",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (confirm != DialogResult.Yes)
                    return;

                try
                {
                    db.RestoreDatabase(dialog.FileName);
                    LoadDevices();
                    MessageBox.Show(
                        "Импорт базы данных завершён успешно.",
                        "Импорт",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Не удалось импортировать базу данных: " + ex.Message,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        void BackupDatabase()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = "Файлы базы данных (*.db)|*.db|Все файлы (*.*)|*.*",
                FileName = $"backup_{DateTime.Now:yyyyMMdd_HHmm}.db"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                try
                {
                    db.BackupDatabase(dialog.FileName);
                    MessageBox.Show(
                        "Резервная копия успешно сохранена.",
                        "Резервное копирование",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Не удалось создать резервную копию: " + ex.Message,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }
    }
}
