using System;
using System.Drawing;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator.Forms
{
    public class DeviceEditForm : Form
    {
        readonly DbService db;
        readonly ComboBox cbType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        readonly TextBox tbName = new TextBox();
        readonly TextBox tbIp = new TextBox();
        readonly TextBox tbMac = new TextBox();
        readonly TextBox tbDescription = new TextBox { Multiline = true, Height = 80 };
        readonly TextBox tbCabinet = new TextBox();

        public Device Result { get; private set; }

        public DeviceEditForm(DbService database, Device source, Cabinet[] cabinets)
        {
            db = database;
            Text = source?.Id > 0 ? "Редактирование устройства" : "Новое устройство";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Width = 440;
            Height = 420;
            MaximizeBox = false;
            MinimizeBox = false;

            cbType.Items.AddRange(new object[]
            {
                DeviceType.Workstation,
                DeviceType.Printer,
                DeviceType.Other
            });
            cbType.FormattingEnabled = true;
            cbType.Format += (s, e) =>
            {
                if (e.ListItem is DeviceType type)
                {
                    e.Value = TranslateType(type);
                }
            };
            cbType.SelectedItem = source?.Type ?? DeviceType.Workstation;

            if (cabinets?.Length > 0)
            {
                var autoComplete = new AutoCompleteStringCollection();
                foreach (var cabinet in cabinets)
                {
                    if (!string.IsNullOrWhiteSpace(cabinet.Name))
                        autoComplete.Add(cabinet.Name);
                }
                tbCabinet.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                tbCabinet.AutoCompleteSource = AutoCompleteSource.CustomSource;
                tbCabinet.AutoCompleteCustomSource = autoComplete;
            }

            if (source != null)
            {
                tbCabinet.Text = source.CabinetName;
                tbName.Text = source.Name;
                tbIp.Text = source.IpAddress;
                tbMac.Text = source.MacAddress;
                tbDescription.Text = source.Description;
            }

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 2,
                AutoSize = true
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));

            AddRow(layout, "Тип", cbType);
            AddRow(layout, "Кабинет", tbCabinet);
            AddRow(layout, "Название", tbName);
            AddRow(layout, "IP адрес", tbIp);
            AddRow(layout, "MAC адрес", tbMac);
            AddRow(layout, "Описание", tbDescription, 80);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            var btnOk = new Button { Text = "Сохранить", DialogResult = DialogResult.None };
            var btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel };
            btnOk.Click += (_, __) => Save(source);
            buttons.Controls.Add(btnOk);
            buttons.Controls.Add(btnCancel);

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.Controls.Add(buttons, 0, layout.RowStyles.Count - 1);
            layout.SetColumnSpan(buttons, 2);

            Controls.Add(layout);
        }

        void AddRow(TableLayoutPanel layout, string label, Control control, int height = 30)
        {
            var rowIndex = layout.RowStyles.Count;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height + 6));
            var lbl = new Label { Text = label, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            control.Dock = DockStyle.Fill;
            layout.Controls.Add(lbl, 0, rowIndex);
            layout.Controls.Add(control, 1, rowIndex);
        }

        void Save(Device original)
        {
            if (cbType.SelectedItem == null)
            {
                MessageBox.Show("Выберите тип устройства.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var cabinetName = tbCabinet.Text.Trim();
            if (string.IsNullOrWhiteSpace(cabinetName))
            {
                MessageBox.Show("Введите кабинет.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(tbName.Text))
            {
                MessageBox.Show("Введите название устройства.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbName.Focus();
                return;
            }
            if (!Validation.IsValidIPv4(tbIp.Text.Trim()))
            {
                MessageBox.Show("Введите корректный IP-адрес.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbIp.Focus();
                return;
            }

            Result = original ?? new Device();
            Result.Type = (DeviceType)cbType.SelectedItem;
            Result.Name = tbName.Text.Trim();
            Result.IpAddress = tbIp.Text.Trim();
            Result.MacAddress = tbMac.Text.Trim();
            Result.Description = tbDescription.Text.Trim();
            try
            {
                Result.CabinetId = db.EnsureCabinet(cabinetName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось сохранить кабинет: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Result.CabinetName = cabinetName;
            Result.AssignedAt = original?.AssignedAt ?? DateTime.Now;

            DialogResult = DialogResult.OK;
            Close();
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
    }
}
