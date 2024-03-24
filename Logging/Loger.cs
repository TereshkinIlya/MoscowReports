using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace Logging
{
    public static class Loger
    {
        private static string LogFolderPath { get; set; }
        private static string LogFilePath { get; set; }
        private static string? TextLine { get; set; }
        static Loger()
        {
            LogFolderPath = Path.Combine(Environment.
                GetFolderPath(Environment.SpecialFolder.Desktop),
                "Приложения А для московской таблицы",
                "Файлы с ошибками");

            LogFilePath = Path.Combine(LogFolderPath, "!Ошибки.txt");
        }
        public static void Log(string message)
        {
            CreateLogFolder();
            try
            {
                File.AppendAllText(LogFilePath, message + "\n");
            }
            catch { }
        }
        public static void Log(string pathFile, string message)
        {
            CreateLogFolder();

            TextLine = $"Имя: {Path.GetFileName(pathFile)} - {message}"+ "\n";
            try
            {
                ExecuteLogging(pathFile, TextLine);
            }
            catch { }
        }
        public static void Log(string pathFile, Excel.Range cell, string message)
        {
            CreateLogFolder();

            TextLine = $"Имя: {Path.GetFileName(pathFile)} - " +
                $"Имя листа: {cell.Worksheet.Name} - " +
                $"Адресс ячейки: {cell.Address} - " +
                $"{message}" + "\n";
            try
            {
                ExecuteLogging(pathFile, TextLine);
            }
            catch { }
        }
        public static void OpenLogFile()
        {
            if(File.Exists(LogFilePath))
            {
                Process.Start("notepad.exe", LogFilePath);
            }
        }
        public static void ClearLogFile()
        {
            if (File.Exists(LogFilePath))
            {
                File.WriteAllText(LogFilePath, string.Empty);
            }
        }
        private static void CreateLogFolder()
        {
            if (!Directory.Exists(LogFolderPath))
                Directory.CreateDirectory(LogFolderPath);
        }
        private static void ExecuteLogging(string pathFile, string text)
        {
            File.Copy(pathFile, Path.Combine(LogFolderPath, Path.GetFileName(pathFile)), true);
            File.AppendAllText(LogFilePath, text);
        }
    }
}
