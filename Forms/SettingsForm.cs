using System;
using System.Drawing;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator.Forms
{
    public class SettingsForm : Form
    {
        readonly TextBox tbPoolStart = new TextBox();
        readonly TextBox tbPoolEnd = new TextBox();
        readonly TextBox tbMask = new TextBox();
        readonly TextBox tbGateway = new TextBox();
        readonly TextBox tbDns1 = new TextBox();
        readonly TextBox tbDns2 = new TextBox();

        public AppSettings Result { get; private set; }

        public SettingsForm(AppSettings current)
        {
            Text = "Настройки сети";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 420;
            Height = 320;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 2,
                RowCount = 7,
                AutoSize = true
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));

            AddRow(layout, "Начало пула", tbPoolStart);
            AddRow(layout, "Конец пула", tbPoolEnd);
            AddRow(layout, "Маска", tbMask);
            AddRow(layout, "Шлюз", tbGateway);
            AddRow(layout, "DNS 1", tbDns1);
            AddRow(layout, "DNS 2", tbDns2);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            var btnOk = new Button { Text = "Сохранить", DialogResult = DialogResult.None };
            var btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel };
            btnOk.Click += (_, __) => Save();
            buttons.Controls.Add(btnOk);
            buttons.Controls.Add(btnCancel);

            layout.Controls.Add(buttons, 0, 6);
            layout.SetColumnSpan(buttons, 2);

            Controls.Add(layout);

            if (current != null)
            {
                tbPoolStart.Text = current.PoolStart;
                tbPoolEnd.Text = current.PoolEnd;
                tbMask.Text = current.Netmask;
                tbGateway.Text = current.Gateway;
                tbDns1.Text = current.Dns1;
                tbDns2.Text = current.Dns2;
            }
        }

        void AddRow(TableLayoutPanel layout, string label, Control control)
        {
            var rowIndex = layout.RowStyles.Count;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            var lbl = new Label { Text = label, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            control.Dock = DockStyle.Fill;
            layout.Controls.Add(lbl, 0, rowIndex);
            layout.Controls.Add(control, 1, rowIndex);
        }

        void Save()
        {
            var settings = new AppSettings
            {
                PoolStart = tbPoolStart.Text.Trim(),
                PoolEnd = tbPoolEnd.Text.Trim(),
                Netmask = tbMask.Text.Trim(),
                Gateway = tbGateway.Text.Trim(),
                Dns1 = tbDns1.Text.Trim(),
                Dns2 = tbDns2.Text.Trim()
            };

            if (!Validate(settings)) return;

            Result = settings;
            DialogResult = DialogResult.OK;
            Close();
        }

        bool Validate(AppSettings settings)
        {
            if (!Validation.IsValidIPv4(settings.PoolStart))
            {
                MessageBox.Show("Введите корректный IP-адрес начала пула.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbPoolStart.Focus();
                return false;
            }
            if (!Validation.IsValidIPv4(settings.PoolEnd))
            {
                MessageBox.Show("Введите корректный IP-адрес конца пула.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbPoolEnd.Focus();
                return false;
            }
            try
            {
                var range = settings.GetPoolRange();
                if (!Validation.IsInRange(range.Start, range.Start, range.End))
                    throw new InvalidOperationException();
            }
            catch
            {
                MessageBox.Show("Начало пула должно быть меньше или равно концу.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbPoolStart.Focus();
                return false;
            }
            if (!Validation.IsValidMask(settings.Netmask))
            {
                MessageBox.Show("Введите корректную маску сети.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbMask.Focus();
                return false;
            }
            if (!Validation.IsValidIPv4(settings.Gateway))
            {
                MessageBox.Show("Введите корректный адрес шлюза.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbGateway.Focus();
                return false;
            }
            if (!Validation.IsValidIPv4(settings.Dns1))
            {
                MessageBox.Show("Введите корректный адрес DNS.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbDns1.Focus();
                return false;
            }
            if (!string.IsNullOrWhiteSpace(settings.Dns2) && !Validation.IsValidIPv4(settings.Dns2))
            {
                MessageBox.Show("Введите корректный адрес второго DNS или оставьте поле пустым.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbDns2.Focus();
                return false;
            }
            return true;
        }
    }
}
