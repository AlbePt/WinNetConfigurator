using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;
using WinNetConfigurator.Utils;
using WinNetConfigurator.UI;

namespace WinNetConfigurator.Forms
{
    public class MainForm : Form
    {
        readonly DbService db;
        readonly ExcelExportService excel;
        readonly NetworkService network;
        readonly BindingList<Device> devices = new BindingList<Device>();
        readonly List<Device> allDevices = new List<Device>();
        readonly DataGridView grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false };
        readonly ContextMenuStrip settingsMenu = new ContextMenuStrip();
        readonly CabinetNameComparer cabinetComparer = new CabinetNameComparer();
        readonly IpAddressComparer ipAddressComparer = new IpAddressComparer();
        readonly NotificationService notificationService;
        Button btnSettings;
        string currentSortProperty;
        bool sortAscending = true;
        AppSettings settings;

        readonly FlowLayoutPanel topButtonsPanel = new FlowLayoutPanel();
        readonly Panel assistantPanel = new Panel();
        readonly Label assistantTitle = new Label();
        readonly Label assistantText = new Label();
        readonly FlowLayoutPanel searchPanel = new FlowLayoutPanel();
        readonly TextBox txtSearch = new TextBox();
        readonly Button btnClearSearch = new Button();
        readonly ToolTip uiToolTip = new ToolTip();
        private Panel notificationsPanel;
        private ListBox notificationsList;
        StatusStrip statusStrip;
        ToolStripStatusLabel statusLabel;

        const string SettingsPassword = "3baRTfg6";

        public MainForm(DbService db, ExcelExportService excelService, NetworkService networkService, AppSettings initialSettings, NotificationService notificationService)
        {
            this.db = db;
            excel = excelService;
            network = networkService;
            settings = initialSettings ?? new AppSettings();
            this.notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));

            KeyPreview = true;

            currentSortProperty = nameof(Device.CabinetName);
            sortAscending = true;

            Text = "WinNetConfigurator";
            Width = 960;
            Height = 600;
            StartPosition = FormStartPosition.CenterScreen;

            BuildLayout();
            UiDefaults.ApplyFormBaseStyle(this);
            ApplyGridStyle(grid);
            LoadDevices();
            notificationService.NotificationsChanged += NotificationServiceOnNotificationsChanged;
            RefreshNotifications();
        }

        void NotificationServiceOnNotificationsChanged(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(RefreshNotifications));
            }
            else
            {
                RefreshNotifications();
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            notificationService.NotificationsChanged -= NotificationServiceOnNotificationsChanged;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            switch (keyData)
            {
                case Keys.Control | Keys.F:
                    txtSearch.Focus();
                    return true;
                case Keys.Delete:
                    DeleteSelected();
                    return true;
                case Keys.Enter:
                    EditSelected();
                    return true;
                default:
                    return base.ProcessCmdKey(ref msg, keyData);
            }
        }

        void BuildLayout()
        {
            BuildTopButtonsPanel();
            BuildAssistantPanel();
            BuildSearchPanel();
            BuildMainGridArea();
            BuildNotificationsPanel();

            statusStrip = UiDefaults.CreateStatusStrip();
            statusLabel = (ToolStripStatusLabel)statusStrip.Items[0];

            Controls.Add(grid);
            Controls.Add(searchPanel);
            Controls.Add(assistantPanel);
            Controls.Add(topButtonsPanel);
            Controls.Add(statusStrip);
        }

        private void BuildNotificationsPanel()
        {
            notificationsPanel = new Panel();
            notificationsPanel.Dock = DockStyle.Right;
            notificationsPanel.Width = 240;
            notificationsPanel.BackColor = AppTheme.SecondaryBackground;
            notificationsPanel.Padding = AppTheme.PaddingNormal;

            var title = new Label();
            title.Text = "Уведомления";
            title.Dock = DockStyle.Top;
            title.Font = new Font(Font, FontStyle.Bold);
            title.Height = 28;

            notificationsList = new ListBox();
            notificationsList.Dock = DockStyle.Fill;

            // пока нет реальных данных — добавим заглушку
            notificationsList.Items.Add("Нет уведомлений");

            notificationsPanel.Controls.Add(notificationsList);
            notificationsPanel.Controls.Add(title);

            // панель должна быть добавлена в Controls формы ДОКУМЕНТАЛЬНО после таблицы или до?
            // Нам нужно, чтобы правая панель была справа от таблицы, значит добавляем в Controls в самом конце:
            this.Controls.Add(notificationsPanel);
        }

        void RefreshNotifications()
        {
            var notifications = notificationService.GetAll();
            notificationsList.BeginUpdate();
            try
            {
                notificationsList.Items.Clear();
                if (notifications == null || notifications.Count == 0)
                {
                    notificationsList.Items.Add("Нет уведомлений");
                    return;
                }

                foreach (var notification in notifications.OrderByDescending(n => n.CreatedAt).Take(20))
                {
                    var timestamp = notification.CreatedAt.ToLocalTime().ToString("dd.MM HH:mm");
                    var text = string.IsNullOrWhiteSpace(notification.Message) ? notification.Title : notification.Message;
                    notificationsList.Items.Add(string.Format("{0} — {1}", timestamp, string.IsNullOrWhiteSpace(text) ? "" : text));
                }
            }
            finally
            {
                notificationsList.EndUpdate();
            }
        }

        private void BuildTopButtonsPanel()
        {
            topButtonsPanel.Dock = DockStyle.Top;
            topButtonsPanel.Height = 48;
            topButtonsPanel.FlowDirection = FlowDirection.LeftToRight;
            topButtonsPanel.Padding = new Padding(12, 10, 12, 6);

            var btnAdd = UiDefaults.CreateTopButton("Добавить", "Добавить новое устройство в базу", uiToolTip);
            UiDefaults.StyleTopCommandButton(btnAdd);
            var btnEdit = UiDefaults.CreateTopButton("Изменить", "Изменить выделенное устройство", uiToolTip);
            UiDefaults.StyleTopCommandButton(btnEdit);
            var btnDelete = UiDefaults.CreateTopButton("Удалить", "Удалить выделенное устройство", uiToolTip);
            UiDefaults.StyleTopCommandButton(btnDelete);
            var btnRefresh = UiDefaults.CreateTopButton("Обновить", "Перечитать список из базы", uiToolTip);
            UiDefaults.StyleTopCommandButton(btnRefresh);
            var btnExport = UiDefaults.CreateTopButton("Экспорт XLSX", "Выгрузить устройства в Excel", uiToolTip);
            UiDefaults.StyleTopCommandButton(btnExport);
            var btnImportXlsx = UiDefaults.CreateTopButton("Импорт XLSX", "Загрузить устройства из Excel-файла", uiToolTip);
            UiDefaults.StyleTopCommandButton(btnImportXlsx);
            btnSettings = UiDefaults.CreateTopButton("Настройки", "Параметры подключения, диапазон IP и т.п.", uiToolTip);
            UiDefaults.StyleTopCommandButton(btnSettings);

            btnAdd.Click += (_, __) => AddDevice();
            btnEdit.Click += (_, __) => EditSelected();
            btnDelete.Click += (_, __) => DeleteSelected();
            btnRefresh.Click += (_, __) => LoadDevices();
            btnExport.Click += (_, __) => Export();
            btnImportXlsx.Click += (_, __) => ImportFromExcel();
            btnSettings.Click += (_, __) => ShowSettingsMenu();

            InitializeSettingsMenu();

            topButtonsPanel.Controls.Add(btnAdd);
            topButtonsPanel.Controls.Add(btnEdit);
            topButtonsPanel.Controls.Add(btnDelete);
            topButtonsPanel.Controls.Add(btnRefresh);
            topButtonsPanel.Controls.Add(btnExport);
            topButtonsPanel.Controls.Add(btnImportXlsx);
            topButtonsPanel.Controls.Add(btnSettings);
        }

        void BuildAssistantPanel()
        {
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

            uiToolTip.SetToolTip(assistantPanel, "Здесь будут появляться подсказки по действиям.");
        }

        private void BuildSearchPanel()
        {
            searchPanel.Dock = DockStyle.Top;
            searchPanel.Height = 36;
            searchPanel.Padding = new Padding(12, 2, 12, 4);
            searchPanel.FlowDirection = FlowDirection.LeftToRight;
            searchPanel.WrapContents = false;

            var lblSearch = new Label
            {
                Text = "Поиск:",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 6, 0)
            };

            txtSearch.Width = 220;
            txtSearch.Margin = new Padding(0, 2, 6, 0);
            txtSearch.TextChanged += (s, e) => ApplyFilter(txtSearch.Text);

            btnClearSearch.Text = "×";
            btnClearSearch.Width = 28;
            btnClearSearch.Height = 24;
            btnClearSearch.Margin = new Padding(0, 2, 0, 0);
            btnClearSearch.Click += (s, e) =>
            {
                txtSearch.Text = "";
                ApplyFilter("");
            };

            searchPanel.Controls.Add(lblSearch);
            searchPanel.Controls.Add(txtSearch);
            searchPanel.Controls.Add(btnClearSearch);
        }

        private void BuildMainGridArea()
        {
            grid.DataSource = devices;

            uiToolTip.SetToolTip(grid, "Двойной щелчок — редактирование. Правой кнопкой — контекстное меню.");

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
        }

        private void ApplyGridStyle(DataGridView grid)
        {
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false;

            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(247, 249, 252);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font(grid.Font, FontStyle.Bold);

            grid.RowHeadersVisible = false;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 252, 255);

            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false;
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

        void EnsureSortGlyph()
        {
            if (string.IsNullOrEmpty(currentSortProperty))
                return;

            var column = grid.Columns.Cast<DataGridViewColumn>()
                .FirstOrDefault(c => c.DataPropertyName == currentSortProperty);
            if (column != null)
                UpdateSortGlyph(column);
        }

        void ShowAssistant(string scenario)
        {
            switch (scenario)
            {
                case "empty":
                    assistantTitle.Text = "Что сделать первым?";
                    assistantText.Text =
                        "Сейчас в базе нет устройств. Нажмите «Добавить устройство» или «Импорт XLSX», " +
                        "чтобы заполнить список.";
                    break;

                case "devices_loaded":
                    assistantTitle.Text = "Подсказка";
                    assistantText.Text =
                        "Двойной щелчок по строке — редактирование. Правой кнопкой — действия. " +
                        "Используйте поиск выше, чтобы быстро найти нужный ПК.";
                    break;

                case "filter_empty":
                    assistantTitle.Text = "Ничего не нашли";
                    assistantText.Text =
                        "Попробуйте ввести часть имени компьютера, кабинета или IP. " +
                        "Кнопка «×» справа от поиска очищает фильтр.";
                    break;

                case "filter_results":
                    assistantTitle.Text = "Фильтр применён";
                    assistantText.Text =
                        "Показаны только устройства, подходящие под запрос. Очистите поиск, чтобы увидеть все.";
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
            allDevices.Clear();

            devices.RaiseListChangedEvents = false;
            devices.Clear();
            foreach (var device in db.GetDevices())
            {
                allDevices.Add(device);
                devices.Add(device);
            }
            devices.RaiseListChangedEvents = true;

            ApplySort(currentSortProperty, preserveDirection: true);
            EnsureSortGlyph();

            ShowAssistant(allDevices.Count == 0 ? "empty" : "devices_loaded");
            SetStatus(allDevices.Count == 0
                ? "Устройства не найдены."
                : $"Загружено устройств: {allDevices.Count}");

            if (!string.IsNullOrWhiteSpace(txtSearch.Text))
                ApplyFilter(txtSearch.Text);
        }

        void SetStatus(string text)
        {
            if (statusLabel != null)
                statusLabel.Text = text;
        }

        void ApplyFilter(string term)
        {
            term = (term ?? string.Empty).Trim();

            if (string.IsNullOrEmpty(term))
            {
                devices.RaiseListChangedEvents = false;
                devices.Clear();
                foreach (var device in allDevices)
                    devices.Add(device);
                devices.RaiseListChangedEvents = true;
                devices.ResetBindings();

                ApplySort(currentSortProperty, preserveDirection: true);
                EnsureSortGlyph();

                ShowAssistant(allDevices.Count == 0 ? "empty" : "devices_loaded");
                SetStatus($"Показаны все устройства ({allDevices.Count})");
                return;
            }

            var lower = term.ToLowerInvariant();

            var filtered = allDevices.Where(d =>
                (!string.IsNullOrEmpty(d.Name) && d.Name.ToLowerInvariant().Contains(lower)) ||
                (!string.IsNullOrEmpty(d.CabinetName) && d.CabinetName.ToLowerInvariant().Contains(lower)) ||
                (!string.IsNullOrEmpty(d.IpAddress) && d.IpAddress.ToLowerInvariant().Contains(lower)) ||
                (!string.IsNullOrEmpty(d.MacAddress) && d.MacAddress.ToLowerInvariant().Contains(lower)) ||
                (!string.IsNullOrEmpty(d.Description) && d.Description.ToLowerInvariant().Contains(lower))
            ).ToList();

            devices.RaiseListChangedEvents = false;
            devices.Clear();
            foreach (var device in filtered)
                devices.Add(device);
            devices.RaiseListChangedEvents = true;
            devices.ResetBindings();

            ApplySort(currentSortProperty, preserveDirection: true);
            EnsureSortGlyph();

            if (filtered.Count == 0)
            {
                ShowAssistant("filter_empty");
                SetStatus($"Ничего не найдено по запросу «{term}»");
            }
            else
            {
                ShowAssistant("filter_results");
                SetStatus($"Найдено: {filtered.Count}");
            }
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
                allDevices.Remove(selected);
                ApplyFilter(txtSearch.Text);
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
