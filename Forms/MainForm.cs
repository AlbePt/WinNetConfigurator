using System;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;

namespace WinNetConfigurator.Forms
{
    public class MainForm : Form
    {
        readonly DbService db;
        readonly ExcelExportService excel;
        readonly BindingList<Device> devices = new BindingList<Device>();
        readonly DataGridView grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false };
        readonly ContextMenuStrip settingsMenu = new ContextMenuStrip();
        Button btnSettings;
        string currentSortProperty;
        bool sortAscending = true;

        const string SettingsPassword = "3baRTfg6";

        public MainForm(DbService db, ExcelExportService excelService)
        {
            this.db = db;
            excel = excelService;

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

            grid.DataSource = devices;
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

            var property = TypeDescriptor.GetProperties(typeof(Device))[propertyName];
            if (property == null)
                return;

            var ordered = sortAscending
                ? devices.OrderBy(d => property.GetValue(d))
                : devices.OrderByDescending(d => property.GetValue(d));

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
            var cabinets = db.GetCabinets().ToArray();
            if (cabinets.Length == 0)
            {
                MessageBox.Show("Сначала добавьте кабинеты в базе (через выбор кабинета при запуске).", "Нет кабинетов", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new DeviceEditForm(null, cabinets))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    if (db.GetDeviceByIp(dialog.Result.IpAddress) != null)
                    {
                        MessageBox.Show("Такой IP уже зарегистрирован в базе.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    dialog.Result.AssignedAt = DateTime.Now;
                    db.InsertDevice(dialog.Result);
                    LoadDevices();
                }
            }
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
            using (var dialog = new DeviceEditForm(new Device
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
