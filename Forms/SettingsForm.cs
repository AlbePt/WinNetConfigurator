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

        readonly Color _errorColor = Color.MistyRose;

        public SettingsForm(AppSettings current)
        {
            Text = "Настройки сети";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 420;
            Height = 380;

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

            AddRow(layout, "Начало пула", tbPoolStart, "Стартовый IP адрес для новых устройств.");
            AddRow(layout, "Конец пула", tbPoolEnd, "Последний IP адрес, доступный для выдачи.");
            AddRow(layout, "Маска", tbMask, "Маска подсети, применяемая к рабочим местам.");
            AddRow(layout, "Шлюз", tbGateway, "IP маршрутизатора по умолчанию.");
            AddRow(layout, "DNS 1", tbDns1, "Основной DNS-сервер для сети.");
            AddRow(layout, "DNS 2", tbDns2, "Резервный DNS (можно оставить пустым).");

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

        void AddRow(TableLayoutPanel layout, string label, TextBox control, string description)
        {
            var rowIndex = layout.RowStyles.Count;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var lbl = new Label { Text = label, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };

            var fieldLayout = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            fieldLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            fieldLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            control.Dock = DockStyle.Fill;
            var hint = new Label
            {
                Text = description,
                AutoSize = true,
                ForeColor = Color.DimGray,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 2, 0, 0)
            };

            fieldLayout.Controls.Add(control, 0, 0);
            fieldLayout.Controls.Add(hint, 0, 1);

            layout.Controls.Add(lbl, 0, rowIndex);
            layout.Controls.Add(fieldLayout, 1, rowIndex);
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
            ResetFieldHighlights();

            if (!ValidateRequiredIp(tbPoolStart, settings.PoolStart, "Введите корректный IP-адрес начала пула."))
                return false;
            if (!ValidateRequiredIp(tbPoolEnd, settings.PoolEnd, "Введите корректный IP-адрес конца пула."))
                return false;
            try
            {
                var range = settings.GetPoolRange();
                if (!Validation.IsInRange(range.Start, range.Start, range.End))
                    throw new InvalidOperationException();
            }
            catch
            {
                Highlight(tbPoolStart);
                Highlight(tbPoolEnd);
                MessageBox.Show("Начало пула должно быть меньше или равно концу.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbPoolStart.Focus();
                return false;
            }
            if (!ValidateRequiredMask(tbMask, settings.Netmask))
                return false;
            if (!ValidateRequiredIp(tbGateway, settings.Gateway, "Введите корректный адрес шлюза."))
                return false;
            if (!ValidateRequiredIp(tbDns1, settings.Dns1, "Введите корректный адрес DNS."))
                return false;
            if (!string.IsNullOrWhiteSpace(settings.Dns2) && !Validation.IsValidIPv4(settings.Dns2))
            {
                Highlight(tbDns2);
                MessageBox.Show("Введите корректный адрес второго DNS или оставьте поле пустым.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbDns2.Focus();
                return false;
            }
            return true;
        }

        void ResetFieldHighlights()
        {
            tbPoolStart.BackColor = SystemColors.Window;
            tbPoolEnd.BackColor = SystemColors.Window;
            tbMask.BackColor = SystemColors.Window;
            tbGateway.BackColor = SystemColors.Window;
            tbDns1.BackColor = SystemColors.Window;
            tbDns2.BackColor = SystemColors.Window;
        }

        bool ValidateRequiredIp(TextBox textBox, string value, string message)
        {
            if (string.IsNullOrWhiteSpace(value) || !Validation.IsValidIPv4(value))
            {
                Highlight(textBox);
                MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox.Focus();
                return false;
            }
            return true;
        }

        bool ValidateRequiredMask(TextBox textBox, string value)
        {
            if (string.IsNullOrWhiteSpace(value) || !Validation.IsValidMask(value))
            {
                Highlight(textBox);
                MessageBox.Show("Введите корректную маску сети.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox.Focus();
                return false;
            }
            return true;
        }

        void Highlight(TextBox textBox)
        {
            textBox.BackColor = _errorColor;
        }
    }
}
