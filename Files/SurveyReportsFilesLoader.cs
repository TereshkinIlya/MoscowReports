using Files.Abstracts;
using Logging;
using System.Security;
namespace Files
{
    public class SurveyReportsFilesLoader : FileManager
    {
        private string RootFolder { get; set; }
        private string SavePath { get; set; }
        private DateTime LimitDate { get; set; }
        private List<FileInfo> Files { get; set; }
        private string[] Years { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public SurveyReportsFilesLoader(DateTime limitDate, string rootFolderPath,
                                string savePathFolder, IProgress<object[]> progress)
        {
            RootFolder = rootFolderPath;
            SavePath = CreateSavingFolder();
            LimitDate = limitDate;
            Progress = progress;
            SavePath = savePathFolder;

            Files = new();
            Years = GetYears();
        }
        /// <summary>
        /// Копирование приложений А на локальный диск
        /// </summary>
        public override void GetInputFiles()
        {
            GetAnnexesAsync();
            CopyAnnexesAsync();
        }
        /// <summary>
        /// Поиск приложений А с учетом предельной даты
        /// </summary>
        /// <returns></returns>
        private void GetAnnexesAsync()
        {
            // Создание входной точки для поиска приложений А
            DirectoryInfo source = new DirectoryInfo(RootFolder);

            var options = new EnumerationOptions()
            {
                IgnoreInaccessible = true,
                RecurseSubdirectories = true
            };
            List<DirectoryInfo> directories = new();

            // Получение значений путей каталогов с учетом года выпуска отчета
            foreach (string year in Years)
            {
                try
                {
                    directories.AddRange(source.GetDirectories("*", options).
                    Where(directory => directory.Name.Contains(year)).ToArray());
                }
                catch (DirectoryNotFoundException)
                {
                    throw;
                }
            }

            // Если выбрана папка на локальном диске (не хранилище ОКУДР)
            if (directories.Count == 0)
                directories.Add(source);

            // Получение значений путей файлов приложений А
            foreach (DirectoryInfo directory in directories)
            {
                Progress.Report(new object[] { 1, 1, "Поиск файлов Приложений А...", 2 });
                try
                {
                    Files.AddRange(directory.GetFiles("*_А.xls*", options).
                    Where(file => file.LastWriteTime.Date > LimitDate.Date
                    &&
                    (file.Extension.EndsWith("xls") ||
                     file.Extension.EndsWith("xlsx") ||
                     file.Extension.EndsWith("xlsm")))
                    .ToArray());
                }
                catch (DirectoryNotFoundException ex)
                {
                    Loger.Log(ex.Message);
                }
                catch (UnauthorizedAccessException ex)
                {
                    Loger.Log(ex.Message);
                }
                catch (SecurityException ex)
                {
                    Loger.Log(ex.Message);
                }
                catch (Exception ex)
                {
                    Loger.Log(ex.Message);
                }
            }
        }
        /// <summary>
        /// Копирование файлов на жесткий диск
        /// </summary>
        /// <returns></returns>
        private void CopyAnnexesAsync()
        {
            int indexProgress = 1;

            foreach (FileInfo file in Files)
            {
                try
                {
                    // Пропуск временных (temp) файлов
                    if (!file.Name.StartsWith("~$"))
                    {
                        File.Copy(file.FullName, Path.Combine(SavePath, file.Name), true);
                    }
                }
                catch (FileNotFoundException ex)
                {
                    Loger.Log(file.FullName, ex.Message);
                }
                catch (UnauthorizedAccessException ex)
                {
                    Loger.Log(file.FullName, ex.Message);
                }
                catch (Exception ex)
                {
                    Loger.Log(file.FullName, ex.Message);
                }
                Progress.Report(new object[] { Files.Count(), indexProgress++, "Копирование файлов Приложений А...", 1 });
            }
        }
        /// <summary>
        /// Годы выпуска отчетов (приложений А)
        /// </summary>
        /// <returns></returns>
        private string[] GetYears()
        {
            string[] years = new string[(DateTime.Now.Year - LimitDate.Year) + 1];
            years[0] = LimitDate.Year.ToString();

            for (int i = 1; i < years.Length; i++)
            {
                years[i] = (LimitDate.Year + i).ToString();
            }

            return years;
        }
        private string CreateSavingFolder()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string savingPath = Path.Combine(desktopPath, "Приложения А для московской таблицы");

            if (!Directory.Exists(savingPath))
                Directory.CreateDirectory(savingPath);
            return savingPath;
        }
    }
}
