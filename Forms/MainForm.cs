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
        readonly ExcelExportService exporter;
        readonly BindingList<Device> devices = new BindingList<Device>();
        readonly DataGridView grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false };

        public MainForm(DbService db, ExcelExportService exporter)
        {
            this.db = db;
            this.exporter = exporter;

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
            var btnClearDb = new Button { Text = "Очистить БД", AutoSize = true };

            btnAdd.Click += (_, __) => AddDevice();
            btnEdit.Click += (_, __) => EditSelected();
            btnDelete.Click += (_, __) => DeleteSelected();
            btnRefresh.Click += (_, __) => LoadDevices();
            btnExport.Click += (_, __) => Export();
            btnClearDb.Click += (_, __) => ClearDatabase();

            panelButtons.Controls.AddRange(new Control[] { btnAdd, btnEdit, btnDelete, btnRefresh, btnExport, btnClearDb });

            grid.DataSource = devices;
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Тип", DataPropertyName = nameof(Device.Type), Width = 120 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Кабинет", DataPropertyName = nameof(Device.CabinetName), Width = 140 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Название", DataPropertyName = nameof(Device.Name), Width = 180 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP адрес", DataPropertyName = nameof(Device.IpAddress), Width = 120 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "MAC адрес", DataPropertyName = nameof(Device.MacAddress), Width = 140 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Описание", DataPropertyName = nameof(Device.Description), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Закреплено", DataPropertyName = nameof(Device.AssignedAt), Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd.MM.yyyy HH:mm" } });
            grid.CellFormatting += Grid_CellFormatting;

            Controls.Add(grid);
            Controls.Add(panelButtons);
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
            if (answer == DialogResult.Yes)
            {
                db.DeleteDevice(selected.Id);
                LoadDevices();
            }
        }

        void Export()
        {
            using (var dialog = new SaveFileDialog { Filter = "Excel файлы (*.xlsx)|*.xlsx", FileName = $"export_{DateTime.Now:yyyyMMdd_HHmm}.xlsx" })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    exporter.ExportDevices(devices, dialog.FileName);
                    MessageBox.Show("Экспорт завершён успешно.", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
    }
}
