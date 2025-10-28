using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator.Forms
{
    public class FreeIpOfferForm : Form
    {
        readonly TextBox tbSelected = new TextBox { Dock = DockStyle.Fill };
        readonly ListBox lbAvailable = new ListBox { Height = 120, Width = 220, Dock = DockStyle.Fill, IntegralHeight = false };
        readonly HashSet<string> pool;
        readonly bool allowCustomEntry;

        public string SelectedIp { get; private set; }

        public FreeIpOfferForm(IEnumerable<string> available, string recommended, string description, bool allowCustomEntry = false)
        {
            pool = new HashSet<string>(available ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            this.allowCustomEntry = allowCustomEntry;

            Text = "Выбор свободного IP";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Width = 420;
            Height = 360;
            MaximizeBox = false;
            MinimizeBox = false;

            foreach (var ip in pool.Take(50))
                lbAvailable.Items.Add(ip);

            tbSelected.Text = recommended;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 6,
                AutoSize = true
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            var lblIntro = new Label
            {
                Text = string.IsNullOrWhiteSpace(description)
                    ? "Выберите свободный IP из пула и закрепите его за рабочим местом."
                    : description,
                Dock = DockStyle.Fill,
                AutoSize = true,
                MaximumSize = new Size(360, 0)
            };

            var lblAvailable = new Label { Text = "Доступные IP (первые 50)", AutoSize = true };
            lbAvailable.SelectedIndexChanged += (_, __) =>
            {
                if (lbAvailable.SelectedItem is string ip)
                    tbSelected.Text = ip;
            };

            var lblSelected = new Label { Text = "Выбранный IP", AutoSize = true };

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            var btnOk = new Button { Text = "Закрепить", DialogResult = DialogResult.None };
            var btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel };
            btnOk.Click += (_, __) => Confirm();
            buttons.Controls.Add(btnOk);
            buttons.Controls.Add(btnCancel);

            layout.Controls.Add(lblIntro, 0, 0);
            layout.Controls.Add(lblAvailable, 0, 1);
            layout.Controls.Add(lbAvailable, 0, 2);
            layout.Controls.Add(lblSelected, 0, 3);
            layout.Controls.Add(tbSelected, 0, 4);
            layout.Controls.Add(buttons, 0, 5);

            Controls.Add(layout);
        }

        void Confirm()
        {
            var candidate = tbSelected.Text.Trim();
            if (string.IsNullOrWhiteSpace(candidate))
            {
                MessageBox.Show("Укажите IP-адрес.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!pool.Contains(candidate))
            {
                if (!allowCustomEntry)
                {
                    MessageBox.Show("Выбранный IP отсутствует среди свободных адресов. Выберите адрес из списка.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!Validation.IsValidIPv4(candidate))
                {
                    MessageBox.Show("Введите корректный IP-адрес.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            SelectedIp = candidate;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
