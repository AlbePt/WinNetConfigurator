using System;
using System.Drawing;
using System.Windows.Forms;
using WinNetConfigurator.UI;

namespace WinNetConfigurator.Forms
{
    public enum RouterIpAction
    {
        None,
        BindOnly,
        BindAndConfigure
    }

    public class RouterIpActionForm : Form
    {
        public RouterIpAction SelectedAction { get; private set; } = RouterIpAction.None;

        public RouterIpActionForm(string currentIp)
        {
            UiDefaults.ApplyFormBaseStyle(this);
            Text = "Выбор действия";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(16);

            var message = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(400, 0),
                Text = string.Format(
                    "У компьютера обнаружен IP-адрес {0} из сети маршрутизатора. Выберите действие.",
                    currentIp)
            };

            var btnBindOnly = new Button
            {
                AutoSize = true,
                Text = "Просто привязать IP"
            };
            btnBindOnly.Click += (_, __) =>
            {
                SelectedAction = RouterIpAction.BindOnly;
                DialogResult = DialogResult.OK;
                Close();
            };

            var btnBindAndConfigure = new Button
            {
                AutoSize = true,
                Text = "Привязать и применить настройки"
            };
            btnBindAndConfigure.Click += (_, __) =>
            {
                SelectedAction = RouterIpAction.BindAndConfigure;
                DialogResult = DialogResult.OK;
                Close();
            };

            var btnCancel = new Button
            {
                AutoSize = true,
                Text = "Отмена",
                DialogResult = DialogResult.Cancel
            };

            CancelButton = btnCancel;
            AcceptButton = btnBindAndConfigure;

            var buttonsPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 12, 0, 0)
            };
            buttonsPanel.Controls.Add(btnBindOnly);
            buttonsPanel.Controls.Add(btnBindAndConfigure);
            buttonsPanel.Controls.Add(btnCancel);

            var layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            layout.Controls.Add(message, 0, 0);
            layout.Controls.Add(buttonsPanel, 0, 1);

            Controls.Add(layout);
        }
    }
}
