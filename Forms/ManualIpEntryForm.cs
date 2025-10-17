using System;
using System.Drawing;
using System.Windows.Forms;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator.Forms
{
    public class ManualIpEntryForm : Form
    {
        readonly TextBox tbIp = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle };
        public string EnteredIp => tbIp.Text.Trim();

        public ManualIpEntryForm(string message, string initialIp)
        {
            Text = "Ввод IP-адреса";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 320);
            Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 204);
            BackColor = Color.FromArgb(244, 246, 251);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 1,
                Padding = new Padding(24),
                BackColor = Color.Transparent
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var card = new TableLayoutPanel
            {
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(24, 24, 24, 18)
            };
            card.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            card.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            card.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            card.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            card.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            card.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            card.Paint += DrawCardBorder;

            var lblTitle = new Label
            {
                Text = "Ввод IP-адреса",
                Font = new Font("Segoe UI", 14F, FontStyle.Bold, GraphicsUnit.Point, 204),
                ForeColor = Color.FromArgb(35, 40, 49),
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 8)
            };

            var lblMessage = new Label
            {
                Text = message,
                AutoSize = true,
                MaximumSize = new Size(420, 0),
                ForeColor = Color.FromArgb(80, 85, 94),
                Margin = new Padding(0, 0, 0, 18)
            };

            var lblInput = new Label
            {
                Text = "IP-адрес",
                AutoSize = true,
                ForeColor = Color.FromArgb(110, 120, 139),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 204),
                Margin = new Padding(0, 0, 0, 6)
            };

            if (!string.IsNullOrWhiteSpace(initialIp) && !initialIp.Trim().StartsWith("192.168.", StringComparison.OrdinalIgnoreCase))
                tbIp.Text = initialIp.Trim();
            tbIp.Margin = new Padding(0, 0, 0, 18);
            tbIp.MinimumSize = new Size(0, 36);
            tbIp.BackColor = Color.FromArgb(248, 249, 252);
            tbIp.Font = new Font("Segoe UI", 11F, FontStyle.Regular, GraphicsUnit.Point, 204);
            tbIp.ForeColor = Color.FromArgb(35, 40, 49);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0),
                WrapContents = false
            };
            var btnOk = CreateDialogButton("Сохранить", true);
            btnOk.DialogResult = DialogResult.None;
            btnOk.Click += (_, __) => Confirm();

            var btnCancel = CreateDialogButton("Отмена", false);
            btnCancel.DialogResult = DialogResult.Cancel;

            buttons.Controls.Add(btnCancel);
            buttons.Controls.Add(btnOk);

            card.Controls.Add(lblTitle, 0, 0);
            card.Controls.Add(lblMessage, 0, 1);
            card.Controls.Add(lblInput, 0, 2);
            card.Controls.Add(tbIp, 0, 3);
            card.Controls.Add(buttons, 0, 4);

            root.Controls.Add(card, 0, 0);
            Controls.Add(root);

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Shown += (_, __) =>
            {
                tbIp.Focus();
                tbIp.SelectAll();
            };
        }

        Button CreateDialogButton(string text, bool primary)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                MinimumSize = new Size(110, 38),
                Padding = new Padding(18, 8, 18, 8),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand,
                BackColor = primary ? Color.FromArgb(71, 99, 255) : Color.White,
                ForeColor = primary ? Color.White : Color.FromArgb(52, 64, 84),
                Font = new Font("Segoe UI", 10F, primary ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Point, 204)
            };
            button.FlatAppearance.BorderSize = primary ? 0 : 1;
            button.FlatAppearance.BorderColor = primary ? Color.FromArgb(71, 99, 255) : Color.FromArgb(208, 213, 222);
            button.FlatAppearance.MouseOverBackColor = primary ? Color.FromArgb(92, 116, 255) : Color.FromArgb(243, 245, 250);
            button.FlatAppearance.MouseDownBackColor = primary ? Color.FromArgb(57, 80, 211) : Color.FromArgb(229, 232, 239);
            return button;
        }

        void DrawCardBorder(object sender, PaintEventArgs e)
        {
            var rect = ((Control)sender).ClientRectangle;
            rect.Width -= 1;
            rect.Height -= 1;
            using var pen = new Pen(Color.FromArgb(221, 225, 233));
            e.Graphics.DrawRectangle(pen, rect);
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
