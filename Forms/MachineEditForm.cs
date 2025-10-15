using System;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator.Forms
{
    public class MachineEditForm : Form
    {
        readonly DbService _db;
        readonly Machine _original;

        ComboBox cbCabinet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 240 };
        TextBox tbHostname = new TextBox { Width = 240 };
        TextBox tbMac = new TextBox { Width = 240 };
        TextBox tbAdapter = new TextBox { Width = 240 };
        TextBox tbIp = new TextBox { Width = 240 };
        CheckBox cbProxyOn = new CheckBox { Text = "Прокси включен", AutoSize = true };
        DateTimePicker dtAssigned = new DateTimePicker { Width = 240, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd HH:mm:ss" };
        ComboBox cbSource = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 240 };

        Button btnOk = new Button { Text = "Сохранить", AutoSize = true };
        Button btnCancel = new Button { Text = "Отмена", AutoSize = true };

        public Machine Result { get; private set; }

        public MachineEditForm(DbService db, Machine machine)
        {
            _db = db;
            _original = machine;

            Text = $"Редактирование записи (ID={machine.Id})";
            StartPosition = FormStartPosition.CenterParent;
            Width = 520; Height = 420;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;

            cbCabinet.Items.AddRange(_db.GetCabinets().Cast<object>().ToArray());
            cbCabinet.Text = machine.CabinetName;

            tbHostname.Text = machine.Hostname;
            tbMac.Text = machine.Mac;
            tbAdapter.Text = machine.AdapterName;
            tbIp.Text = machine.Ip;
            cbProxyOn.Checked = machine.ProxyOn;
            dtAssigned.Value = machine.AssignedAt == default ? DateTime.Now : machine.AssignedAt;

            cbSource.Items.AddRange(new[] { "auto", "inventory" });
            cbSource.SelectedItem = (machine.Source == "inventory") ? "inventory" : "auto";

            BuildLayout();

            btnOk.Click += (s, e) => OnSave();
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
        }

        void BuildLayout()
        {
            var tl = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Padding = new Padding(12)
            };
            tl.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            void AddRow(string label, Control input)
            {
                int r = tl.RowCount; tl.RowCount++;
                tl.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var lbl = new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
                input.Margin = new Padding(0, 3, 0, 3);
                input.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                tl.Controls.Add(lbl, 0, r);
                tl.Controls.Add(input, 1, r);
            }

            AddRow("Кабинет:", cbCabinet);
            AddRow("Hostname:", tbHostname);
            AddRow("MAC:", tbMac);
            AddRow("Адаптер:", tbAdapter);
            AddRow("IP:", tbIp);

            int rProxy = tl.RowCount; tl.RowCount++;
            tl.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbProxyOn.Margin = new Padding(0, 6, 0, 6);
            tl.Controls.Add(cbProxyOn, 1, rProxy);

            AddRow("Дата назначения:", dtAssigned);
            AddRow("Источник:", cbSource);

            var btnCol = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                WrapContents = false,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 12, 0, 0)
            };
            btnCol.Controls.Add(btnOk);
            btnCol.Controls.Add(btnCancel);

            int rBtn = tl.RowCount; tl.RowCount++;
            tl.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tl.Controls.Add(btnCol, 1, rBtn);

            Controls.Add(tl);
        }

        void OnSave()
        {
            string cabinetName = cbCabinet.Text?.Trim();
            if (string.IsNullOrWhiteSpace(cabinetName)) { MessageBox.Show("Укажите кабинет."); return; }
            if (string.IsNullOrWhiteSpace(tbHostname.Text)) { MessageBox.Show("Укажите Hostname."); return; }
            if (string.IsNullOrWhiteSpace(tbMac.Text)) { MessageBox.Show("Укажите MAC."); return; }
            if (string.IsNullOrWhiteSpace(tbAdapter.Text)) { MessageBox.Show("Укажите имя адаптера."); return; }
            if (!Validation.IsValidIPv4(tbIp.Text)) { MessageBox.Show("Некорректный IP."); return; }

            Result = new Machine
            {
                Id = _original.Id,
                CabinetName = cabinetName,
                Hostname = tbHostname.Text.Trim(),
                Mac = tbMac.Text.Trim(),
                AdapterName = tbAdapter.Text.Trim(),
                Ip = tbIp.Text.Trim(),
                ProxyOn = cbProxyOn.Checked,
                AssignedAt = dtAssigned.Value,
                Source = (cbSource.SelectedItem?.ToString() ?? "auto")
            };

            try
            {
                _db.UpdateMachine(Result.Id, Result.CabinetId, Result.Ip);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения: " + ex.Message);
            }
        }
    }
}


