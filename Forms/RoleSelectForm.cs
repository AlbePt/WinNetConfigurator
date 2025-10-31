using System;
using System.Drawing;
using System.Windows.Forms;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Forms
{
    public class RoleSelectForm : Form
    {
        public UserRole SelectedRole { get; private set; } = UserRole.Operator;

        public RoleSelectForm()
        {
            Text = "Выбор режима";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(360, 180);

            var lbl = new Label
            {
                Text = "Кто сейчас работает с программой?",
                Dock = DockStyle.Top,
                Height = 36,
                TextAlign = ContentAlignment.MiddleCenter
            };

            var btnAssistant = new Button
            {
                Text = "Ассистент (обход кабинетов)",
                AutoSize = true,
                Width = 260
            };
            btnAssistant.Click += (_, __) =>
            {
                SelectedRole = UserRole.Operator;   // будем считать, что ассистент = оператор
                DialogResult = DialogResult.OK;
                Close();
            };

            var btnAdmin = new Button
            {
                Text = "Администратор",
                AutoSize = true,
                Width = 260
            };
            btnAdmin.Click += (_, __) =>
            {
                SelectedRole = UserRole.Administrator;
                DialogResult = DialogResult.OK;
                Close();
            };

            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                Padding = new Padding(20),
                WrapContents = false
            };
            panel.Controls.Add(lbl);
            panel.Controls.Add(btnAssistant);
            panel.Controls.Add(btnAdmin);

            Controls.Add(panel);
        }
    }
}
