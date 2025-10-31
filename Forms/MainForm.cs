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
        readonly NotificationService notifications;
        readonly TaskQueueService taskQueue;
        readonly BindingList<Device> devices = new BindingList<Device>();
        readonly List<Device> allDevices = new List<Device>();
        readonly DataGridView grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false };
        readonly CabinetNameComparer cabinetComparer = new CabinetNameComparer();
        readonly IpAddressComparer ipAddressComparer = new IpAddressComparer();
        readonly ToolTip adminToolTip = new ToolTip { ShowAlways = true };
        SplitContainer mainSplit;
        TabControl adminTabControl;
        TabPage adminTabPage;
        Label adminStatusLabel;
        Button btnAdminLogin;
        Button btnNetworkSettings;
        Button btnImportDb;
        Button btnBackupDb;
        Button btnClearDb;
        ListView notificationsList;
        ListView tasksList;
        ProgressBar taskProgressBar;
        Label taskProgressLabel;
        TextBox filterTextBox;
        StatusStrip statusStrip;
        ToolStripStatusLabel devicesStatusLabel;
        ToolStripStatusLabel tasksStatusLabel;
        Button btnSettings;
        string currentSortProperty;
        bool sortAscending = true;
        string currentFilterText = string.Empty;
        AppSettings settings;

        const string SettingsPassword = "3baRTfg6";
        const int MaxPasswordAttempts = 3;
        bool _adminMode;
        int _failedAdminAttempts;
        bool _adminLocked;

        public MainForm(DbService db, ExcelExportService excelService, NetworkService networkService, AppSettings initialSettings, NotificationService notificationService, TaskQueueService taskQueueService)
        {
            this.db = db;
            excel = excelService;
            network = networkService;
            notifications = notificationService;
            taskQueue = taskQueueService;
            settings = initialSettings ?? new AppSettings();

            currentSortProperty = nameof(Device.CabinetName);
            sortAscending = true;

            Text = "WinNetConfigurator";
            Width = 960;
            Height = 600;
            StartPosition = FormStartPosition.CenterScreen;

            BuildLayout();
            LoadDevices();

            if (notifications != null)
                notifications.NotificationsChanged += Notifications_NotificationsChanged;
            if (taskQueue != null)
                taskQueue.TasksChanged += TaskQueue_TasksChanged;

            RefreshNotificationsList();
            RefreshTasksList();
        }

        void BuildLayout()
        {
            mainSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                FixedPanel = FixedPanel.Panel2,
                SplitterWidth = 6,
                Panel2MinSize = 260
            };
            mainSplit.SplitterDistance = Width - 300;

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
            btnSettings = new Button { Text = "Администратор", AutoSize = true };

            btnAdd.Click += (_, __) => AddDevice();
            btnEdit.Click += (_, __) => EditSelected();
            btnDelete.Click += (_, __) => DeleteSelected();
            btnRefresh.Click += (_, __) => LoadDevices();
            btnExport.Click += (_, __) => Export();
            btnImportXlsx.Click += (_, __) => ImportFromExcel();
            btnSettings.Click += (_, __) => ActivateAdminPanel();

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
            grid.CellDoubleClick += Grid_CellDoubleClick;

            var filterPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 36,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(12, 6, 12, 0),
                AutoSize = true,
                WrapContents = false
            };
            var filterLabel = new Label { Text = "Фильтр:", AutoSize = true, Anchor = AnchorStyles.Left };
            filterTextBox = new TextBox { Width = 220, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            filterTextBox.TextChanged += (_, __) => ApplyFilter();
            filterPanel.Controls.Add(filterLabel);
            filterPanel.Controls.Add(filterTextBox);

            var leftPanel = new Panel { Dock = DockStyle.Fill };
            leftPanel.Controls.Add(panelButtons);
            leftPanel.Controls.Add(filterPanel);
            leftPanel.Controls.Add(grid);

            mainSplit.Panel1.Controls.Add(leftPanel);

            BuildAdminPanel();
            mainSplit.Panel2.Controls.Add(adminTabControl);

            statusStrip = new StatusStrip();
            devicesStatusLabel = new ToolStripStatusLabel("Устройств: 0");
            tasksStatusLabel = new ToolStripStatusLabel(string.Empty);
            statusStrip.Items.Add(devicesStatusLabel);
            statusStrip.Items.Add(new ToolStripStatusLabel { Spring = true });
            statusStrip.Items.Add(tasksStatusLabel);

            Controls.Add(mainSplit);
            Controls.Add(statusStrip);
        }

        void BuildAdminPanel()
        {
            adminTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(12),
                Padding = new Point(12, 6)
            };

            adminTabPage = new TabPage("Администрирование") { Padding = new Padding(12) };
            var adminLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(8),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1
            };
            adminLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            adminStatusLabel = new Label
            {
                AutoSize = true,
                Text = "Администратор: нужен пароль",
                Font = new Font(Font, FontStyle.Bold)
            };

            btnAdminLogin = new Button { Text = "Войти как администратор...", AutoSize = true };
            btnAdminLogin.Click += (_, __) => RequestAdminAccess();

            btnNetworkSettings = new Button { Text = "Настройки сети...", AutoSize = true };
            btnImportDb = new Button { Text = "Импорт БД...", AutoSize = true };
            btnBackupDb = new Button { Text = "Резервная копия БД...", AutoSize = true };
            btnClearDb = new Button { Text = "Очистить БД...", AutoSize = true };

            btnNetworkSettings.Click += (_, __) => ExecuteAdminAction(EditSettings);
            btnImportDb.Click += (_, __) => ExecuteAdminAction(ImportDatabase);
            btnBackupDb.Click += (_, __) => ExecuteAdminAction(BackupDatabase);
            btnClearDb.Click += (_, __) => ExecuteAdminAction(ClearDatabase);

            adminToolTip.SetToolTip(btnNetworkSettings, "Нужен админ-доступ");
            adminToolTip.SetToolTip(btnImportDb, "Нужен админ-доступ");
            adminToolTip.SetToolTip(btnBackupDb, "Нужен админ-доступ");
            adminToolTip.SetToolTip(btnClearDb, "Нужен админ-доступ");

            var buttonsLayout = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true
            };
            buttonsLayout.Controls.Add(btnNetworkSettings);
            buttonsLayout.Controls.Add(btnImportDb);
            buttonsLayout.Controls.Add(CreateWarningLabel());
            buttonsLayout.Controls.Add(btnBackupDb);
            buttonsLayout.Controls.Add(btnClearDb);
            buttonsLayout.Controls.Add(CreateWarningLabel());

            adminLayout.Controls.Add(adminStatusLabel);
            adminLayout.Controls.Add(btnAdminLogin);
            adminLayout.Controls.Add(buttonsLayout);

            adminTabPage.Controls.Add(adminLayout);

            var notificationsTab = new TabPage("Уведомления") { Padding = new Padding(12) };
            notificationsList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            notificationsList.Columns.Add("Время", 120);
            notificationsList.Columns.Add("Уровень", 100);
            notificationsList.Columns.Add("Заголовок", 200);
            notificationsTab.Controls.Add(notificationsList);

            var tasksTab = new TabPage("Задачи") { Padding = new Padding(12) };
            var tasksLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            tasksLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tasksLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tasksLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            taskProgressLabel = new Label { AutoSize = true, Visible = false };
            taskProgressBar = new ProgressBar { Dock = DockStyle.Top, Height = 18, Visible = false };

            tasksList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            tasksList.Columns.Add("Название", 180);
            tasksList.Columns.Add("Статус", 120);
            tasksList.Columns.Add("Прогресс", 80);
            tasksList.Columns.Add("Запущена", 140);

            tasksLayout.Controls.Add(taskProgressLabel, 0, 0);
            tasksLayout.Controls.Add(taskProgressBar, 0, 1);
            tasksLayout.Controls.Add(tasksList, 0, 2);

            tasksTab.Controls.Add(tasksLayout);

            adminTabControl.TabPages.Add(adminTabPage);
            adminTabControl.TabPages.Add(notificationsTab);
            adminTabControl.TabPages.Add(tasksTab);

            UpdateAdminControls();
        }

        Label CreateWarningLabel()
        {
            return new Label
            {
                Text = "Осторожно: действие необратимо",
                AutoSize = true,
                ForeColor = Color.DarkRed,
                Margin = new Padding(3, 3, 3, 6)
            };
        }

        void ActivateAdminPanel()
        {
            if (adminTabControl == null || adminTabPage == null)
                return;

            adminTabControl.SelectedTab = adminTabPage;
            adminTabControl.Focus();
        }

        void ExecuteAdminAction(Action action)
        {
            if (!_adminMode)
            {
                ShowStatusMessage("Нужен админ-доступ");
                return;
            }

            action?.Invoke();
        }

        void RequestAdminAccess()
        {
            if (_adminLocked)
                return;

            using (var dialog = new PasswordPromptForm())
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                if (!string.Equals(dialog.EnteredPassword, SettingsPassword, StringComparison.Ordinal))
                {
                    _failedAdminAttempts++;
                    if (_failedAdminAttempts >= MaxPasswordAttempts)
                    {
                        _adminLocked = true;
                        ShowStatusMessage("Админ-доступ заблокирован");
                    }
                    else
                    {
                        ShowStatusMessage("Неверный пароль");
                    }
                    UpdateAdminControls();
                    return;
                }
            }

            _adminMode = true;
            _failedAdminAttempts = 0;
            ShowStatusMessage("Доступ администратора получен");
            UpdateAdminControls();
        }

        void UpdateAdminControls()
        {
            var hasAccess = _adminMode && !_adminLocked;
            adminStatusLabel.Text = hasAccess
                ? "Администратор: доступ получен"
                : _adminLocked ? "Администратор: доступ заблокирован" : "Администратор: нужен пароль";
            btnAdminLogin.Enabled = !_adminMode && !_adminLocked;
            SetAdminButtonState(btnNetworkSettings, hasAccess);
            SetAdminButtonState(btnImportDb, hasAccess);
            SetAdminButtonState(btnBackupDb, hasAccess);
            SetAdminButtonState(btnClearDb, hasAccess);
        }

        void SetAdminButtonState(Button button, bool enabled)
        {
            if (button == null)
                return;

            button.Enabled = enabled;
        }

        void Notifications_NotificationsChanged(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => Notifications_NotificationsChanged(sender, e)));
                return;
            }

            RefreshNotificationsList();
        }

        void TaskQueue_TasksChanged(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => TaskQueue_TasksChanged(sender, e)));
                return;
            }

            RefreshTasksList();
        }

        void RefreshNotificationsList()
        {
            if (notificationsList == null || notifications == null)
                return;

            notificationsList.BeginUpdate();
            notificationsList.Items.Clear();
            foreach (var notification in notifications.GetForUser(null).OrderByDescending(n => n.CreatedAt))
            {
                var item = new ListViewItem(notification.CreatedAt.ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss"))
                {
                    Tag = notification
                };
                item.SubItems.Add(notification.Severity.ToString());
                item.SubItems.Add(notification.Title ?? string.Empty);
                ApplyNotificationStyle(item, notification.Severity);
                notificationsList.Items.Add(item);
            }
            notificationsList.EndUpdate();
        }

        void ApplyNotificationStyle(ListViewItem item, NotificationSeverity severity)
        {
            switch (severity)
            {
                case NotificationSeverity.Info:
                    item.BackColor = Color.White;
                    item.ForeColor = Color.Black;
                    break;
                case NotificationSeverity.Warning:
                    item.BackColor = Color.FromArgb(255, 252, 229);
                    item.ForeColor = Color.DarkOrange;
                    break;
                case NotificationSeverity.Error:
                    item.BackColor = Color.MistyRose;
                    item.ForeColor = Color.DarkRed;
                    break;
                case NotificationSeverity.Critical:
                    item.BackColor = Color.IndianRed;
                    item.ForeColor = Color.White;
                    break;
            }
        }

        void RefreshTasksList()
        {
            if (tasksList == null || taskQueue == null)
                return;

            tasksList.BeginUpdate();
            tasksList.Items.Clear();
            var tasks = taskQueue.ListTasks().OrderByDescending(t => t.CreatedAt).ToList();
            foreach (var task in tasks)
            {
                var item = new ListViewItem(task.Name)
                {
                    Tag = task
                };
                item.SubItems.Add(GetTaskStatusText(task));
                item.SubItems.Add(task.Progress + "%");
                item.SubItems.Add(task.CreatedAt.ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss"));

                if (task.Status == BackgroundTaskStatus.Failed)
                {
                    item.BackColor = Color.MistyRose;
                    item.ForeColor = Color.DarkRed;
                }

                tasksList.Items.Add(item);
            }
            tasksList.EndUpdate();

            var runningTask = tasks.FirstOrDefault(t => t.Status == BackgroundTaskStatus.Running);
            if (runningTask != null)
            {
                taskProgressLabel.Text = "Выполняется... " + (string.IsNullOrWhiteSpace(runningTask.StateMessage) ? runningTask.Name : runningTask.StateMessage);
                taskProgressLabel.Visible = true;
                taskProgressBar.Value = Math.Max(0, Math.Min(100, runningTask.Progress));
                taskProgressBar.Visible = true;
            }
            else
            {
                taskProgressLabel.Visible = false;
                taskProgressBar.Visible = false;
            }

            UpdateStatusBar();
        }

        string GetTaskStatusText(BackgroundTask task)
        {
            switch (task.Status)
            {
                case BackgroundTaskStatus.Pending:
                    return "В очереди";
                case BackgroundTaskStatus.Running:
                    return "Выполняется";
                case BackgroundTaskStatus.Completed:
                    return string.IsNullOrEmpty(task.ResultMessage) ? "Завершена" : task.ResultMessage;
                case BackgroundTaskStatus.Failed:
                    return "Ошибка";
                case BackgroundTaskStatus.Cancelled:
                    return "Отменена";
                default:
                    return task.Status.ToString();
            }
        }

        void ShowStatusMessage(string message)
        {
            if (tasksStatusLabel == null)
                return;

            tasksStatusLabel.Text = message;
            tasksStatusLabel.ForeColor = SystemColors.ControlText;
        }

        void UpdateStatusBar()
        {
            if (devicesStatusLabel != null)
                devicesStatusLabel.Text = "Устройств: " + devices.Count;

            if (tasksStatusLabel != null && taskQueue != null)
            {
                var activeCount = taskQueue.ListTasks().Count(t => t.Status == BackgroundTaskStatus.Pending || t.Status == BackgroundTaskStatus.Running);
                if (activeCount > 0)
                {
                    tasksStatusLabel.Text = "Задачи: " + activeCount;
                    tasksStatusLabel.ForeColor = Color.DarkOrange;
                }
                else if (string.IsNullOrEmpty(tasksStatusLabel.Text) || tasksStatusLabel.Text.StartsWith("Задачи:"))
                {
                    tasksStatusLabel.Text = string.Empty;
                    tasksStatusLabel.ForeColor = SystemColors.ControlText;
                }
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            if (notifications != null)
                notifications.NotificationsChanged -= Notifications_NotificationsChanged;
            if (taskQueue != null)
                taskQueue.TasksChanged -= TaskQueue_TasksChanged;
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

        void LoadDevices()
        {
            allDevices.Clear();
            allDevices.AddRange(db.GetDevices());
            RefreshDeviceView();
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
            RefreshDeviceView();
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

        void RefreshDeviceView()
        {
            IEnumerable<Device> query = allDevices;

            if (!string.IsNullOrWhiteSpace(currentFilterText))
            {
                var filter = currentFilterText.Trim();
                query = query.Where(d =>
                    (!string.IsNullOrEmpty(d.CabinetName) && d.CabinetName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (!string.IsNullOrEmpty(d.Name) && d.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (!string.IsNullOrEmpty(d.IpAddress) && d.IpAddress.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0));
            }

            var list = ApplyOrdering(query).ToList();

            devices.RaiseListChangedEvents = false;
            devices.Clear();
            foreach (var device in list)
                devices.Add(device);
            devices.RaiseListChangedEvents = true;
            devices.ResetBindings();

            if (!string.IsNullOrEmpty(currentSortProperty))
            {
                var column = grid.Columns.Cast<DataGridViewColumn>()
                    .FirstOrDefault(c => c.DataPropertyName == currentSortProperty);
                if (column != null)
                    UpdateSortGlyph(column);
            }

            UpdateStatusBar();
        }

        IEnumerable<Device> ApplyOrdering(IEnumerable<Device> items)
        {
            if (string.IsNullOrEmpty(currentSortProperty))
                return items;

            switch (currentSortProperty)
            {
                case nameof(Device.CabinetName):
                    return sortAscending
                        ? items.OrderBy(d => d.CabinetName, cabinetComparer)
                        : items.OrderByDescending(d => d.CabinetName, cabinetComparer);
                case nameof(Device.IpAddress):
                    return sortAscending
                        ? items.OrderBy(d => d.IpAddress, ipAddressComparer)
                        : items.OrderByDescending(d => d.IpAddress, ipAddressComparer);
                default:
                    var property = TypeDescriptor.GetProperties(typeof(Device))[currentSortProperty];
                    if (property == null)
                        return items;
                    return sortAscending
                        ? items.OrderBy(d => property.GetValue(d))
                        : items.OrderByDescending(d => property.GetValue(d));
            }
        }

        void ApplyFilter()
        {
            currentFilterText = filterTextBox?.Text ?? string.Empty;
            RefreshDeviceView();
        }

        void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                EditSelected();
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

                    ShowStatusMessage("Импорт из XLSX выполнен успешно");
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
                    ShowStatusMessage("Настройки обновлены");
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
            ShowStatusMessage("База данных очищена");
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
                    ShowStatusMessage("Импорт базы данных завершён успешно");
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
                FileName = $"backup_{DateTime.Now:yyyyMMddHHmmss}.db"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                try
                {
                    db.BackupDatabase(dialog.FileName);
                    ShowStatusMessage("Резервная копия создана");
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
