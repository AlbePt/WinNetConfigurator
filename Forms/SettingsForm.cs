using System;
using System.Drawing;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;
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

        readonly DbService db;
        readonly Action onDatabaseChanged;

        public AppSettings Result { get; private set; }

        public SettingsForm(AppSettings current, DbService dbService, Action onDatabaseChanged)
        {
            db = dbService ?? throw new ArgumentNullException(nameof(dbService));
            this.onDatabaseChanged = onDatabaseChanged;

            Text = "Настройки";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 540;
            Height = 520;
            Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 204);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16, 16, 16, 8),
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var tabs = new TabControl { Dock = DockStyle.Fill };
            tabs.TabPages.Add(CreateNetworkTab(current));
            tabs.TabPages.Add(CreateDatabaseTab());
            root.Controls.Add(tabs, 0, 0);

            var buttonsPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(0, 16, 0, 0)
            };

            var btnSave = CreatePrimaryButton("Сохранить");
            btnSave.DialogResult = DialogResult.None;
            btnSave.Click += (_, __) => SaveNetworkSettings();

            var btnCancel = CreateSecondaryButton("Отмена");
            btnCancel.DialogResult = DialogResult.Cancel;

            buttonsPanel.Controls.Add(btnSave);
            buttonsPanel.Controls.Add(btnCancel);

            root.Controls.Add(buttonsPanel, 0, 1);

            Controls.Add(root);

            AcceptButton = btnSave;
            CancelButton = btnCancel;

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

        TabPage CreateNetworkTab(AppSettings current)
        {
            var page = new TabPage("Сеть") { Padding = new Padding(16) };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));

            var lblInfo = new Label
            {
                Text = "Укажите параметры пула IP-адресов и основные сетевые настройки.",
                AutoSize = true,
                MaximumSize = new Size(440, 0),
                ForeColor = Color.FromArgb(90, 98, 118),
                Margin = new Padding(0, 0, 0, 12)
            };
            layout.Controls.Add(lblInfo, 0, 0);
            layout.SetColumnSpan(lblInfo, 2);

            AddRow(layout, "Начало пула", tbPoolStart, 1);
            AddRow(layout, "Конец пула", tbPoolEnd, 2);
            AddRow(layout, "Маска", tbMask, 3);
            AddRow(layout, "Шлюз", tbGateway, 4);
            AddRow(layout, "DNS 1", tbDns1, 5);
            AddRow(layout, "DNS 2", tbDns2, 6);

            page.Controls.Add(layout);
            return page;
        }

        TabPage CreateDatabaseTab()
        {
            var page = new TabPage("База данных") { Padding = new Padding(16) };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lblIntro = new Label
            {
                Text = "Инструменты ниже позволяют управлять базой данных приложения. Операции необратимы — перед импортом или очисткой рекомендуется сделать резервную копию.",
                AutoSize = true,
                MaximumSize = new Size(440, 0),
                ForeColor = Color.FromArgb(90, 98, 118),
                Margin = new Padding(0, 0, 0, 16)
            };
            layout.Controls.Add(lblIntro, 0, 0);

            var toolsPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                Margin = new Padding(0, 0, 0, 12)
            };

            var btnImportDb = CreateSecondaryButton("Импорт БД");
            btnImportDb.Click += (_, __) => ImportDatabase();

            var btnBackupDb = CreateSecondaryButton("Резервная копия БД");
            btnBackupDb.Click += (_, __) => BackupDatabase();

            var btnClearDb = CreateDangerButton("Очистить БД");
            btnClearDb.Click += (_, __) => ClearDatabase();

            toolsPanel.Controls.Add(btnImportDb);
            toolsPanel.Controls.Add(btnBackupDb);
            toolsPanel.Controls.Add(btnClearDb);

            layout.Controls.Add(toolsPanel, 0, 1);

            page.Controls.Add(layout);
            return page;
        }

        Button CreatePrimaryButton(string text)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                Padding = new Padding(18, 8, 18, 8),
                Margin = new Padding(0, 0, 0, 0),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                BackColor = Color.FromArgb(71, 99, 255),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(92, 116, 255);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(57, 80, 211);
            return button;
        }

        Button CreateSecondaryButton(string text)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                Padding = new Padding(18, 8, 18, 8),
                Margin = new Padding(8, 0, 0, 0),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                BackColor = Color.FromArgb(229, 232, 239),
                ForeColor = Color.FromArgb(35, 40, 49),
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(214, 218, 226);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(198, 203, 213);
            return button;
        }

        Button CreateDangerButton(string text)
        {
            var button = CreateSecondaryButton(text);
            button.BackColor = Color.FromArgb(255, 231, 231);
            button.ForeColor = Color.FromArgb(176, 24, 44);
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 217, 217);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(242, 200, 200);
            return button;
        }

        void AddRow(TableLayoutPanel layout, string label, Control control, int rowIndex)
        {
            while (layout.RowStyles.Count <= rowIndex)
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lbl = new Label
            {
                Text = label,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 12, 6)
            };
            control.Margin = new Padding(0, 4, 0, 4);
            control.Dock = DockStyle.Fill;

            layout.Controls.Add(lbl, 0, rowIndex);
            layout.Controls.Add(control, 1, rowIndex);
        }

        void SaveNetworkSettings()
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

            if (!Validate(settings))
                return;

            Result = settings;
            DialogResult = DialogResult.OK;
            Close();
        }

        bool Validate(AppSettings settings)
        {
            if (!Validation.IsValidIPv4(settings.PoolStart))
            {
                MessageBox.Show(this,
                    "Введите корректный IP-адрес начала пула.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                tbPoolStart.Focus();
                return false;
            }
            if (!Validation.IsValidIPv4(settings.PoolEnd))
            {
                MessageBox.Show(this,
                    "Введите корректный IP-адрес конца пула.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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
                MessageBox.Show(this,
                    "Начало пула должно быть меньше или равно концу.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                tbPoolStart.Focus();
                return false;
            }
            if (!Validation.IsValidMask(settings.Netmask))
            {
                MessageBox.Show(this,
                    "Введите корректную маску сети.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                tbMask.Focus();
                return false;
            }
            if (!Validation.IsValidIPv4(settings.Gateway))
            {
                MessageBox.Show(this,
                    "Введите корректный адрес шлюза.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                tbGateway.Focus();
                return false;
            }
            if (!Validation.IsValidIPv4(settings.Dns1))
            {
                MessageBox.Show(this,
                    "Введите корректный адрес DNS.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                tbDns1.Focus();
                return false;
            }
            if (!string.IsNullOrWhiteSpace(settings.Dns2) && !Validation.IsValidIPv4(settings.Dns2))
            {
                MessageBox.Show(this,
                    "Введите корректный адрес второго DNS или оставьте поле пустым.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                tbDns2.Focus();
                return false;
            }
            return true;
        }

        void ClearDatabase()
        {
            var answer = MessageBox.Show(this,
                "Это удалит все устройства, кабинеты и настройки сети. Продолжить?",
                "Очистка базы данных",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            if (answer != DialogResult.Yes)
                return;

            db.ClearDatabase();
            onDatabaseChanged?.Invoke();

            MessageBox.Show(this,
                "База данных очищена.",
                "Готово",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        void ImportDatabase()
        {
            using (var dialog = new OpenFileDialog { Filter = "Файлы базы данных (*.db)|*.db|Все файлы (*.*)|*.*" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                var confirm = MessageBox.Show(this,
                    "Текущая база данных будет заменена выбранным файлом. Продолжить?",
                    "Импорт базы данных",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (confirm != DialogResult.Yes)
                    return;

                try
                {
                    db.RestoreDatabase(dialog.FileName);
                    onDatabaseChanged?.Invoke();
                    MessageBox.Show(this,
                        "Импорт базы данных завершён успешно.",
                        "Импорт",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this,
                        "Не удалось импортировать базу данных: " + ex.Message,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        void BackupDatabase()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = "Файлы базы данных (*.db)|*.db|Все файлы (*.*)|*.*",
                FileName = $"backup_{DateTime.Now:yyyyMMdd_HHmm}.db"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                try
                {
                    db.BackupDatabase(dialog.FileName);
                    MessageBox.Show(this,
                        "Резервная копия успешно сохранена.",
                        "Резервное копирование",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this,
                        "Не удалось создать резервную копию: " + ex.Message,
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }
    }
}
