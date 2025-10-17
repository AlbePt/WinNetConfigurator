using System.Drawing;
using System.Windows.Forms;

namespace WinNetConfigurator.Forms
{
    public class PasswordPromptForm : Form
    {
        readonly TextBox tbPassword = new TextBox { UseSystemPasswordChar = true, Dock = DockStyle.Fill };
        public string EnteredPassword => tbPassword.Text.Trim();

        public PasswordPromptForm()
        {
            Text = "Пароль";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 360;
            Height = 180;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lblPrompt = new Label
            {
                Text = "Введите пароль для доступа к настройкам:",
                AutoSize = true,
                Dock = DockStyle.Fill,
                MaximumSize = new Size(320, 0)
            };

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };

            var btnOk = new Button { Text = "ОК", DialogResult = DialogResult.OK, AutoSize = true };
            var btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, AutoSize = true };

            buttons.Controls.Add(btnOk);
            buttons.Controls.Add(btnCancel);

            layout.Controls.Add(lblPrompt, 0, 0);
            layout.Controls.Add(tbPassword, 0, 1);
            layout.Controls.Add(buttons, 0, 2);

            Controls.Add(layout);

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Shown += (_, __) =>
            {
                tbPassword.Focus();
                tbPassword.SelectAll();
            };
        }
    }
}
