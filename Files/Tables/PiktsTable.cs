using Excel = Microsoft.Office.Interop.Excel;

namespace Files.Tables
{
    public class PiktsTable
    {
        private string Path { get; set; }
        internal object[,] TableData { get; set; }

        internal Dictionary<string, List<int>> PiktsTableColumns = new Dictionary<string, List<int>>()
        {
            {"ID ППМТ",new() },
            {"Состояние нитки",new() },
            {"Характеристика труб руслового участка",new() },
            {"Дата прогона",new() },
            {"Предельная дата эксплуатации трубной секции с дефектом",new() },
            {"Местоположение дефекта (русло/пойма)",new() },
            {"Дата окончания срока безопасной эксплуатации",new() },
            {"Номер заключения по оценке технического состояния",new() },
            {"Дата выдачи заключения по оценке технического состояния",new() },
            {"Дата окончания срока безопасной эксплуатации по параметрам",new() },
            {"Организация, проводившая экспертную оценку",new() },
            {"Рекомендации по приведению ПМТ в нормативное состояние",new() },
            {"Признак ремонта",new() },
            {"Метод ремонта",new() }
        };
        public PiktsTable(DataAnalyzer dataAnalyzer, string path)
        {
            ArgumentNullException.ThrowIfNull(dataAnalyzer);
            ArgumentNullException.ThrowIfNull(path, "Выберите файл с таблицей дефектов из ПИКТС");

            Path = path;
            if (!File.Exists(Path)) throw new ArgumentException("Файл c выгрузкой из ПИКТС не существует");

            dataAnalyzer.GetHeaderColumns(Path, PiktsTableColumns, "31. Отчет ПАО");

            Excel.Workbook workbook = ExcelApp.Run.Workbooks.Open(Path);
            Excel.Worksheet worksheet = workbook.Worksheets.Item["31. Отчет ПАО"];
            TableData = worksheet.Range["A1", worksheet.UsedRange].Value;
            workbook.Close();
        }
    }
}
