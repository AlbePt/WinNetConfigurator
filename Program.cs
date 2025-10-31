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

            var db = new DbService();
            var exporter = new ExcelExportService(db);
            var network = new NetworkService();
            var settings = db.LoadSettings();

            if (settings == null || !settings.IsComplete())
            {
                using (var settingsForm = new SettingsForm(settings ?? new AppSettings()))
                {
                    if (settingsForm.ShowDialog() != DialogResult.OK)
                        return;
                    settings = settingsForm.Result;
                    db.SaveSettings(settings);
                }
            }

            using (var roleForm = new RoleSelectForm())
            {
                if (roleForm.ShowDialog() != DialogResult.OK)
                    return;

                if (roleForm.SelectedRole == UserRole.Operator)
                {
                    Application.Run(new AssistantForm(db, network, settings));
                }
                else
                {
                    Application.Run(new MainForm(db, exporter, network, settings));
                }
            }
        }
    }
}
