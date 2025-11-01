using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.UI;

namespace WinNetConfigurator.Forms
{
    public class CabinetSelectionForm : Form
    {
        readonly TextBox tbCabinet = new TextBox { Width = 220 };
        public string CabinetName { get; private set; } = string.Empty;

        public CabinetSelectionForm(IEnumerable<Cabinet> cabinets)
        {
            UiDefaults.ApplyFormBaseStyle(this);
            Text = "Выбор кабинета";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 420;
            Height = 220;

            var autoComplete = new AutoCompleteStringCollection();
            foreach (var cab in cabinets.OrderBy(c => c.Name))
            {
                if (!string.IsNullOrWhiteSpace(cab.Name))
                    autoComplete.Add(cab.Name);
            }
            tbCabinet.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            tbCabinet.AutoCompleteSource = AutoCompleteSource.CustomSource;
            tbCabinet.AutoCompleteCustomSource = autoComplete;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 2,
                RowCount = 2,
                AutoSize = true
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));

            AddRow(layout, "Кабинет", tbCabinet);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            var btnOk = new Button { Text = "Выбрать", DialogResult = DialogResult.None };
            var btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel };
            btnOk.Click += (_, __) => Confirm();
            buttons.Controls.Add(btnOk);
            buttons.Controls.Add(btnCancel);

            layout.Controls.Add(buttons, 0, 1);
            layout.SetColumnSpan(buttons, 2);

            Controls.Add(layout);
        }

        void AddRow(TableLayoutPanel layout, string label, Control control)
        {
            var rowIndex = layout.RowStyles.Count;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            var lbl = new Label { Text = label, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            control.Dock = DockStyle.Fill;
            layout.Controls.Add(lbl, 0, rowIndex);
            layout.Controls.Add(control, 1, rowIndex);
        }

        void Confirm()
        {
            if (!string.IsNullOrWhiteSpace(tbCabinet.Text))
            {
                CabinetName = tbCabinet.Text.Trim();
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            MessageBox.Show("Укажите кабинет.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
