using Files;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace MoscowReports
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.

            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(UnhandledExceptions);
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());

        }
        public static void UnhandledExceptions(object sender, UnhandledExceptionEventArgs e)
        {
            int id;

            GetWindowThreadProcessId(ExcelApp.Run.Hwnd, out id);
            Process process = Process.GetProcessById(id);
            process.Kill();

            DialogResult result =  MessageBox.Show((e.ExceptionObject as Exception)?.Message);

            Application.Restart();
            Environment.Exit(Environment.ExitCode);

        }
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
    }
}