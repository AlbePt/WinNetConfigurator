using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Forms
{
    public class CabinetSelectionForm : Form
    {
        readonly ComboBox cbCabinets = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        readonly TextBox tbNewCabinet = new TextBox { Width = 220 };

        public Cabinet SelectedCabinet { get; private set; }

        public CabinetSelectionForm(IEnumerable<Cabinet> cabinets)
        {
            Text = "Выбор кабинета";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 420;
            Height = 220;

            foreach (var cab in cabinets.OrderBy(c => c.Name))
                cbCabinets.Items.Add(cab);
            if (cbCabinets.Items.Count > 0)
                cbCabinets.SelectedIndex = 0;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 2,
                RowCount = 3,
                AutoSize = true
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));

            AddRow(layout, "Существующий кабинет", cbCabinets);
            AddRow(layout, "Новый кабинет", tbNewCabinet);

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

            layout.Controls.Add(buttons, 0, 2);
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
            if (!string.IsNullOrWhiteSpace(tbNewCabinet.Text))
            {
                SelectedCabinet = new Cabinet { Name = tbNewCabinet.Text.Trim() };
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            if (cbCabinets.SelectedItem is Cabinet cabinet)
            {
                SelectedCabinet = cabinet;
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            MessageBox.Show("Выберите существующий кабинет или укажите новый.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
