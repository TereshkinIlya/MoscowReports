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
        private List<string> Files { get; set; }
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
            GetAnnexes();
            CopyAnnexes();
        }
        /// <summary>
        /// Поиск приложений А с учетом предельной даты
        /// </summary>
        private void GetAnnexes()
        {
            var options = new EnumerationOptions()
            {
                IgnoreInaccessible = true,
                RecurseSubdirectories = true
            };

            List<string> directories = new();

            foreach (string year in Years)
            {
                Progress.Report(new object[] { 1, 1, "Поиск каталогов с Приложениями А...", 2 });
                try
                {
                    directories.AddRange(Directory.GetDirectories(RootFolder, "*", options).
                        Where(directory => directory.Contains(year)).ToArray());
                }
                catch (DirectoryNotFoundException ex)
                {
                    Loger.Log(ex.Message);
                }
            }

            // Если выбрана папка на локальном диске (не хранилище ОКУДР)
            if (directories.Count == 0)
                directories.Add(RootFolder);

            foreach (string directory in directories)
            {
                Progress.Report(new object[] { 1, 1, "Поиск файлов Приложений А...", 2 });
                try
                {
                    Files.AddRange(Directory.GetFiles(directory, "*_А.xls*", options).
                    Where(file => File.GetLastWriteTime(file) > LimitDate.Date
                    &&
                    (Path.GetExtension(file).EndsWith("xls") ||
                    Path.GetExtension(file).EndsWith("xlsx") ||
                    Path.GetExtension(file).EndsWith("xlsm"))
                    ).ToArray());
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
        private void CopyAnnexes()
        {
            int indexProgress = 1;

            foreach (string file in Files)
            {
                try
                {
                    // Пропуск временных (temp) файлов
                    if (!Path.GetFileName(file).StartsWith("~$"))
                    {
                        File.Copy(file, Path.Combine(SavePath, Path.GetFileName(file)), true);
                    }
                }
                catch (FileNotFoundException ex)
                {
                    Loger.Log(Path.GetFileName(file), ex.Message);
                }
                catch (UnauthorizedAccessException ex)
                {
                    Loger.Log(Path.GetFileName(file), ex.Message);
                }
                catch (Exception ex)
                {
                    Loger.Log(Path.GetFileName(file), ex.Message);
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
