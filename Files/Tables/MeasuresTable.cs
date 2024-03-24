namespace Files.Tables
{
    public class MeasuresTable
    {
        internal string Path { get; set; }

        internal Dictionary<string, List<int>> TableColumns = new Dictionary<string, List<int>>()
        {
            {"№№ п/п",new() },
            {"ОСТ",new() },
            {"WM",new() },
            {"MFL (CDC)",new() },
            {"CDL",new() },
            {"Мероприятия",new() },
            {"Номер заключения",new() },
            {"Дата выдачи заключения",new() },
            {"ВТД",new() },
            {"ЭХЗ",new() }
        };

        public MeasuresTable(DataAnalyzer dataAnalyzer, string? path)
        {
            ArgumentNullException.ThrowIfNull(dataAnalyzer);
            ArgumentNullException.ThrowIfNull(path, "Выберите файл с таблицей мероприятий");

            Path = path;
            if (!File.Exists(Path)) throw new ArgumentException("файл Таблицы мероприятий не существует");

            dataAnalyzer.GetHeaderColumns(Path, TableColumns, "Сводная");
        }
    }
}
