using System;
using System.Drawing;
using System.Windows.Forms;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator.Forms
{
    public class ManualIpEntryForm : Form
    {
        readonly TextBox tbIp = new TextBox { Dock = DockStyle.Fill };
        public string EnteredIp => tbIp.Text.Trim();

        public ManualIpEntryForm(string message, string initialIp)
        {
            Text = "Ввод IP-адреса";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 420;
            Height = 240;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 4,
                AutoSize = true
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lblMessage = new Label
            {
                Text = message,
                AutoSize = true,
                Dock = DockStyle.Fill,
                MaximumSize = new Size(360, 0)
            };

            var lblInput = new Label
            {
                Text = "IP-адрес",
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 12, 0, 4)
            };

            tbIp.Text = initialIp ?? string.Empty;

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            var btnOk = new Button { Text = "Сохранить", DialogResult = DialogResult.None, AutoSize = true };
            var btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, AutoSize = true };
            btnOk.Click += (_, __) => Confirm();
            buttons.Controls.Add(btnOk);
            buttons.Controls.Add(btnCancel);

            layout.Controls.Add(lblMessage, 0, 0);
            layout.Controls.Add(lblInput, 0, 1);
            layout.Controls.Add(tbIp, 0, 2);
            layout.Controls.Add(buttons, 0, 3);

            Controls.Add(layout);

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Shown += (_, __) =>
            {
                tbIp.Focus();
                tbIp.SelectAll();
            };
        }

        void Confirm()
        {
            var candidate = tbIp.Text.Trim();
            if (!Validation.IsValidIPv4(candidate))
            {
                MessageBox.Show("Введите корректный IP-адрес.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbIp.Focus();
                tbIp.SelectAll();
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
