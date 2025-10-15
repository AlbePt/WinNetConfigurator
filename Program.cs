using System;
using System.IO;
using System.Windows.Forms;

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
            Directory.SetCurrentDirectory(baseDir);

            Application.ThreadException += (s, e) =>
            {
                Utils.Logger.Log("UI ThreadException: " + e.Exception);
                MessageBox.Show(e.Exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                Utils.Logger.Log("UnhandledException: " + e.ExceptionObject);
            };

            Application.Run(new Forms.MainForm());
        }
    }
}
