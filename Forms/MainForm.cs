using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;

namespace WinNetConfigurator.Forms
{
    public class MainForm : Form
    {
        const string SettingsPassword = "3baRTfg6";

        readonly DbService db;
        readonly ExcelExportService excel;
        readonly List<Device> allDevices = new List<Device>();
        readonly BindingList<Device> filteredDevices = new BindingList<Device>();
        readonly BindingSource bindingSource = new BindingSource();
        readonly DataGridView grid;
        readonly ComboBox cbFilterType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
        readonly Dictionary<string, TextBox> textFilters = new Dictionary<string, TextBox>();
        readonly Label lblSummary = new Label
        {
            Dock = DockStyle.Bottom,
            Height = 28,
            TextAlign = ContentAlignment.MiddleRight,
            Padding = new Padding(0, 0, 16, 0),
            ForeColor = Color.FromArgb(110, 120, 139)
        };
        readonly Dictionary<string, Func<Device, object>> sortSelectors;
        string currentSortProperty;
        bool sortAscending;

        public MainForm(DbService db, ExcelExportService excelService)
        {
            this.db = db;
            excel = excelService;

            Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 204);
            Text = "WinNetConfigurator";
            Width = 1100;
            Height = 680;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(248, 249, 252);

            sortSelectors = new Dictionary<string, Func<Device, object>>
            {
                [nameof(Device.Type)] = d => d.Type,
                [nameof(Device.CabinetName)] = d => d.CabinetName ?? string.Empty,
                [nameof(Device.Name)] = d => d.Name ?? string.Empty,
                [nameof(Device.IpAddress)] = d => d.IpAddress ?? string.Empty,
                [nameof(Device.MacAddress)] = d => d.MacAddress ?? string.Empty,
                [nameof(Device.Description)] = d => d.Description ?? string.Empty,
                [nameof(Device.AssignedAt)] = d => d.AssignedAt
            };
            currentSortProperty = nameof(Device.AssignedAt);
            sortAscending = false;

            grid = BuildGrid();
            bindingSource.DataSource = filteredDevices;
            grid.DataSource = bindingSource;

            BuildLayout();
            LoadDevices();
        }

        DataGridView BuildGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AutoGenerateColumns = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                GridColor = Color.FromArgb(230, 233, 239),
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            };

            dgv.ColumnHeadersHeight = 44;
            dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(238, 240, 245),
                ForeColor = Color.FromArgb(52, 64, 84),
                Font = new Font(Font, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleLeft,
                Padding = new Padding(12, 10, 12, 10)
            };
            dgv.DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.White,
                ForeColor = Color.FromArgb(35, 40, 49),
                SelectionBackColor = Color.FromArgb(216, 227, 255),
                SelectionForeColor = Color.FromArgb(35, 40, 49),
                Padding = new Padding(8, 6, 8, 6)
            };
            dgv.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(246, 248, 252),
                SelectionBackColor = Color.FromArgb(216, 227, 255),
                SelectionForeColor = Color.FromArgb(35, 40, 49),
                Padding = new Padding(8, 6, 8, 6)
            };
            dgv.EnableHeadersVisualStyles = false;

            dgv.Columns.Add(CreateTextColumn("Тип", nameof(Device.Type), 140));
            dgv.Columns.Add(CreateTextColumn("Кабинет", nameof(Device.CabinetName), 160));
            dgv.Columns.Add(CreateTextColumn("Название", nameof(Device.Name), 220));
            dgv.Columns.Add(CreateTextColumn("IP адрес", nameof(Device.IpAddress), 140));
            dgv.Columns.Add(CreateTextColumn("MAC адрес", nameof(Device.MacAddress), 170));
            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Описание",
                DataPropertyName = nameof(Device.Description),
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 220,
                SortMode = DataGridViewColumnSortMode.Programmatic
            });
            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Закреплено",
                DataPropertyName = nameof(Device.AssignedAt),
                Width = 170,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "dd.MM.yyyy HH:mm" },
                SortMode = DataGridViewColumnSortMode.Programmatic
            });

            dgv.CellFormatting += Grid_CellFormatting;
            dgv.ColumnHeaderMouseClick += Grid_ColumnHeaderMouseClick;

            var doubleBufferedProperty = typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance);
            doubleBufferedProperty?.SetValue(dgv, true, null);

            return dgv;
        }

        DataGridViewTextBoxColumn CreateTextColumn(string header, string property, int width)
        {
            return new DataGridViewTextBoxColumn
            {
                HeaderText = header,
                DataPropertyName = property,
                Width = width,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
        }

        void BuildLayout()
        {
            var topContainer = new Panel
            {
                Dock = DockStyle.Top,
                Padding = new Padding(20, 18, 20, 12),
                BackColor = Color.FromArgb(242, 244, 248),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lblTitle = new Label
            {
                Text = "WinNetConfigurator",
                Font = new Font("Segoe UI", 20F, FontStyle.Bold, GraphicsUnit.Point, 204),
                ForeColor = Color.FromArgb(35, 40, 49),
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 12)
            };
            layout.Controls.Add(lblTitle, 0, 0);

            var commandPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = true,
                Margin = new Padding(0, 0, 0, 12)
            };
            var btnAdd = CreateActionButton("Добавить устройство", true);
            var btnEdit = CreateActionButton("Редактировать");
            var btnDelete = CreateActionButton("Удалить");
            var btnRefresh = CreateActionButton("Обновить");
            var btnExport = CreateActionButton("Экспорт в XLSX");
            var btnImportXlsx = CreateActionButton("Импорт из XLSX");
            var btnSettings = CreateActionButton("Настройки");

            btnAdd.Click += (_, __) => AddDevice();
            btnEdit.Click += (_, __) => EditSelected();
            btnDelete.Click += (_, __) => DeleteSelected();
            btnRefresh.Click += (_, __) => LoadDevices();
            btnExport.Click += (_, __) => Export();
            btnImportXlsx.Click += (_, __) => ImportFromExcel();
            btnSettings.Click += (_, __) => EditSettings();

            commandPanel.Controls.AddRange(new Control[]
            {
                btnAdd,
                btnEdit,
                btnDelete,
                btnRefresh,
                btnExport,
                btnImportXlsx,
                btnSettings
            });
            layout.Controls.Add(commandPanel, 0, 1);

            var filtersPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = true
            };
            BuildFilters(filtersPanel);
            layout.Controls.Add(filtersPanel, 0, 2);

            topContainer.Controls.Add(layout);

            Controls.Add(grid);
            Controls.Add(lblSummary);
            Controls.Add(topContainer);
        }

        Button CreateActionButton(string text, bool primary = false)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                Margin = new Padding(0, 0, 12, 12),
                Padding = new Padding(18, 8, 18, 8),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand,
                BackColor = primary ? Color.FromArgb(71, 99, 255) : Color.FromArgb(229, 232, 239),
                ForeColor = primary ? Color.White : Color.FromArgb(35, 40, 49)
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = primary ? Color.FromArgb(92, 116, 255) : Color.FromArgb(214, 218, 226);
            button.FlatAppearance.MouseDownBackColor = primary ? Color.FromArgb(57, 80, 211) : Color.FromArgb(198, 203, 213);
            return button;
        }

        void BuildFilters(FlowLayoutPanel panel)
        {
            var typeContainer = CreateFilterContainer("Тип");
            cbFilterType.Items.Add(new DeviceTypeFilterOption("Все типы", null));
            cbFilterType.Items.Add(new DeviceTypeFilterOption("Рабочее место", DeviceType.Workstation));
            cbFilterType.Items.Add(new DeviceTypeFilterOption("Принтер", DeviceType.Printer));
            cbFilterType.Items.Add(new DeviceTypeFilterOption("Другое устройство", DeviceType.Other));
            cbFilterType.SelectedIndex = 0;
            cbFilterType.SelectedIndexChanged += (_, __) => ApplyFiltersAndSorting();
            typeContainer.Controls.Add(cbFilterType);
            panel.Controls.Add(typeContainer);

            AddTextFilter(panel, "Кабинет", nameof(Device.CabinetName));
            AddTextFilter(panel, "Название", nameof(Device.Name));
            AddTextFilter(panel, "IP адрес", nameof(Device.IpAddress));
            AddTextFilter(panel, "MAC адрес", nameof(Device.MacAddress));
            AddTextFilter(panel, "Описание", nameof(Device.Description));
            AddTextFilter(panel, "Закреплено", nameof(Device.AssignedAt));
        }

        FlowLayoutPanel CreateFilterContainer(string caption)
        {
            var container = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 0, 16, 12)
            };
            container.Controls.Add(new Label
            {
                Text = caption,
                AutoSize = true,
                ForeColor = Color.FromArgb(110, 120, 139),
                Margin = new Padding(0, 0, 0, 4)
            });
            return container;
        }

        void AddTextFilter(FlowLayoutPanel panel, string caption, string property)
        {
            var container = CreateFilterContainer(caption);
            var tb = new TextBox
            {
                Width = 180
            };
            tb.TextChanged += (_, __) => ApplyFiltersAndSorting();
            container.Controls.Add(tb);
            panel.Controls.Add(container);
            textFilters[property] = tb;
        }

        void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var column = grid.Columns[e.ColumnIndex];
            if (column.DataPropertyName == nameof(Device.Type) && e.Value is DeviceType type)
            {
                e.Value = TranslateType(type);
                e.FormattingApplied = true;
            }
            else if (column.DataPropertyName == nameof(Device.AssignedAt) && e.Value is DateTime date)
            {
                e.Value = date.ToString("dd.MM.yyyy HH:mm");
                e.FormattingApplied = true;
            }
        }

        void Grid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var column = grid.Columns[e.ColumnIndex];
            if (string.IsNullOrWhiteSpace(column.DataPropertyName))
                return;

            if (currentSortProperty == column.DataPropertyName)
                sortAscending = !sortAscending;
            else
            {
                currentSortProperty = column.DataPropertyName;
                sortAscending = true;
            }

            ApplyFiltersAndSorting();
            UpdateSortGlyphs(column);
        }

        void UpdateSortGlyphs(DataGridViewColumn activeColumn)
        {
            foreach (DataGridViewColumn column in grid.Columns)
            {
                column.HeaderCell.SortGlyphDirection = SortOrder.None;
            }

            activeColumn.HeaderCell.SortGlyphDirection = sortAscending ? SortOrder.Ascending : SortOrder.Descending;
        }

        void LoadDevices()
        {
            allDevices.Clear();
            allDevices.AddRange(db.GetDevices());
            ApplyFiltersAndSorting();
        }

        void ApplyFiltersAndSorting()
        {
            IEnumerable<Device> query = allDevices;

            if (cbFilterType.SelectedItem is DeviceTypeFilterOption option && option.Type.HasValue)
                query = query.Where(d => d.Type == option.Type.Value);

            foreach (var filter in textFilters)
            {
                var text = filter.Value.Text.Trim();
                if (string.IsNullOrEmpty(text))
                    continue;

                var comparison = StringComparison.OrdinalIgnoreCase;
                if (filter.Key == nameof(Device.AssignedAt))
                {
                    query = query.Where(d => d.AssignedAt.ToString("dd.MM.yyyy HH:mm").IndexOf(text, comparison) >= 0);
                }
                else
                {
                    query = query.Where(d =>
                    {
                        var value = GetPropertyValue(d, filter.Key);
                        return value.IndexOf(text, comparison) >= 0;
                    });
                }
            }

            if (!string.IsNullOrEmpty(currentSortProperty) && sortSelectors.TryGetValue(currentSortProperty, out var selector))
            {
                query = sortAscending
                    ? query.OrderBy(selector)
                    : query.OrderByDescending(selector);
            }
            else
            {
                query = query.OrderByDescending(d => d.AssignedAt);
            }

            var result = query.ToList();

            filteredDevices.RaiseListChangedEvents = false;
            filteredDevices.Clear();
            foreach (var device in result)
                filteredDevices.Add(device);
            filteredDevices.RaiseListChangedEvents = true;

            bindingSource.ResetBindings(false);
            lblSummary.Text = $"Показано {filteredDevices.Count} из {allDevices.Count}";

            if (!string.IsNullOrEmpty(currentSortProperty))
            {
                var column = grid.Columns.Cast<DataGridViewColumn>()
                    .FirstOrDefault(c => c.DataPropertyName == currentSortProperty);
                if (column != null)
                    UpdateSortGlyphs(column);
            }
        }

        string GetPropertyValue(Device device, string property)
        {
            switch (property)
            {
                case nameof(Device.CabinetName):
                    return device.CabinetName ?? string.Empty;
                case nameof(Device.Name):
                    return device.Name ?? string.Empty;
                case nameof(Device.IpAddress):
                    return device.IpAddress ?? string.Empty;
                case nameof(Device.MacAddress):
                    return device.MacAddress ?? string.Empty;
                case nameof(Device.Description):
                    return device.Description ?? string.Empty;
                default:
                    return string.Empty;
            }
        }

        Device GetSelectedDevice()
        {
            return grid.CurrentRow?.DataBoundItem as Device;
        }

        void AddDevice()
        {
            var cabinets = db.GetCabinets().ToArray();
            if (cabinets.Length == 0)
            {
                MessageBox.Show(this,
                    "Сначала добавьте кабинеты в базе (через выбор кабинета при запуске).",
                    "Нет кабинетов",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new DeviceEditForm(null, cabinets))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    if (db.GetDeviceByIp(dialog.Result.IpAddress) != null)
                    {
                        MessageBox.Show(this,
                            "Такой IP уже зарегистрирован в базе.",
                            "Ошибка",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
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
                MessageBox.Show(this,
                    "Выберите устройство в таблице.",
                    "Нет выбора",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
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
                        MessageBox.Show(this,
                            "Такой IP уже зарегистрирован в базе.",
                            "Ошибка",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
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
                MessageBox.Show(this,
                    "Выберите устройство в таблице.",
                    "Нет выбора",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var answer = MessageBox.Show(this,
                $"Удалить устройство {selected.Name} ({selected.IpAddress})?",
                "Подтверждение",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (answer != DialogResult.Yes)
                return;

            if (db.DeleteDevice(selected.Id))
            {
                LoadDevices();
            }
            else
            {
                MessageBox.Show(this,
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
                    excel.ExportDevices(filteredDevices, dialog.FileName);
                    MessageBox.Show(this,
                        "Экспорт завершён успешно.",
                        "Экспорт",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }

        void ImportFromExcel()
        {
            using (var dialog = new OpenFileDialog { Filter = "Excel файлы (*.xlsx)|*.xlsx" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                var confirm = MessageBox.Show(this,
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
                        MessageBox.Show(this,
                            "В выбранном файле не найдено устройств для импорта.",
                            "Импорт из XLSX",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }

                    db.ReplaceDevices(imported);
                    LoadDevices();

                    MessageBox.Show(this,
                        "Импорт из XLSX выполнен успешно.",
                        "Импорт из XLSX",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this,
                        "Не удалось выполнить импорт из XLSX: " + ex.Message,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        void EditSettings()
        {
            using (var passwordPrompt = new PasswordPromptForm(SettingsPassword))
            {
                if (passwordPrompt.ShowDialog(this) != DialogResult.OK)
                    return;
            }

            var current = db.LoadSettings();
            using (var dialog = new SettingsForm(current, db, LoadDevices))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    db.SaveSettings(dialog.Result);
                    MessageBox.Show(this,
                        "Настройки обновлены.",
                        "Настройки",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }

        string TranslateType(DeviceType type)
        {
            switch (type)
            {
                case DeviceType.Printer:
                    return "Принтер";
                case DeviceType.Other:
                    return "Другое устройство";
                default:
                    return "Рабочее место";
            }
        }

        class DeviceTypeFilterOption
        {
            public DeviceTypeFilterOption(string caption, DeviceType? type)
            {
                Caption = caption;
                Type = type;
            }

            public string Caption { get; }

            public DeviceType? Type { get; }

            public override string ToString() => Caption;
        }
    }
}
