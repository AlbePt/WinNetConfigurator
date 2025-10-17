using System;
using System.Drawing;
using System.Windows.Forms;

namespace WinNetConfigurator.Forms
{
    public class PasswordPromptForm : Form
    {
        readonly string password;
        readonly TextBox tbPassword = new TextBox { UseSystemPasswordChar = true, Width = 240 };

        public PasswordPromptForm(string password)
        {
            this.password = password ?? throw new ArgumentNullException(nameof(password));

            Text = "Введите пароль";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 380;
            Height = 210;
            Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 204);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                ColumnCount = 1,
                RowCount = 4,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lblDescription = new Label
            {
                Text = "Для открытия меню настроек требуется пароль.",
                AutoSize = true,
                MaximumSize = new Size(320, 0),
                ForeColor = Color.FromArgb(90, 98, 118)
            };

            var lblPassword = new Label
            {
                Text = "Пароль",
                AutoSize = true,
                Margin = new Padding(0, 12, 0, 4)
            };

            tbPassword.Margin = new Padding(0, 0, 0, 6);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            var btnOk = new Button
            {
                Text = "Подтвердить",
                DialogResult = DialogResult.None,
                AutoSize = true,
                Padding = new Padding(18, 8, 18, 8),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                BackColor = Color.FromArgb(71, 99, 255),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnOk.FlatAppearance.BorderSize = 0;
            btnOk.FlatAppearance.MouseOverBackColor = Color.FromArgb(92, 116, 255);
            btnOk.FlatAppearance.MouseDownBackColor = Color.FromArgb(57, 80, 211);
            btnOk.Click += (_, __) => Confirm();

            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                AutoSize = true,
                Padding = new Padding(18, 8, 18, 8),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                BackColor = Color.FromArgb(229, 232, 239),
                ForeColor = Color.FromArgb(35, 40, 49),
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.FlatAppearance.MouseOverBackColor = Color.FromArgb(214, 218, 226);
            btnCancel.FlatAppearance.MouseDownBackColor = Color.FromArgb(198, 203, 213);

            buttons.Controls.Add(btnOk);
            buttons.Controls.Add(btnCancel);

            layout.Controls.Add(lblDescription, 0, 0);
            layout.Controls.Add(lblPassword, 0, 1);
            layout.Controls.Add(tbPassword, 0, 2);
            layout.Controls.Add(buttons, 0, 3);

            Controls.Add(layout);

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Shown += (_, __) =>
            {
                tbPassword.Focus();
            };
        }

        void Confirm()
        {
            if (string.Equals(tbPassword.Text, password))
            {
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            MessageBox.Show(this,
                "Неверный пароль.",
                "Доступ запрещён",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            tbPassword.Focus();
            tbPassword.SelectAll();
        }
    }
}
