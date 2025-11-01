using System;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            System.IO.Directory.SetCurrentDirectory(baseDir);

            Application.ThreadException += (s, e) =>
            {
                Logger.Log("UI ThreadException: " + e.Exception);
                MessageBox.Show(e.Exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                Logger.Log("UnhandledException: " + e.ExceptionObject);
            };

            var db = new DbService();
            var network = new NetworkService();
            var exporter = new ExcelExportService();
            var notificationService = new NotificationService();

            var settings = db.LoadSettings();
            if (!settings.IsComplete())
            {
                using (var settingsForm = new SettingsForm(settings))
                {
                    if (settingsForm.ShowDialog() != DialogResult.OK)
                        return;
                    settings = settingsForm.Result;
                    db.SaveSettings(settings);
                }
            }

            Application.Run(new MainForm(db, exporter, network, settings, notificationService));
        }
    }
}
