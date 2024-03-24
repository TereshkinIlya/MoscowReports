namespace Files.Tables
{
    public class AcumulativeTable
    {
        internal string Path { get; set; }

        internal Dictionary<string, List<int>> TableColumns = new Dictionary<string, List<int>>()
        {
            {"ID",new() },
            {"Дата выдачи заключения",new() },
            {"Номер ОТС",new() },
            {"Проведение ВТД",new() },
            {"Обследование корр",new() }
        };
        public AcumulativeTable(DataAnalyzer dataAnalyzer, string? path)
        {
            ArgumentNullException.ThrowIfNull(dataAnalyzer);
            ArgumentNullException.ThrowIfNull(path, "Выберите файл с накопительной таблицей");

            Path = path;
            if (!File.Exists(Path)) throw new ArgumentException("файл Накопительной таблицы не существует");

            dataAnalyzer.GetHeaderColumns(Path, TableColumns, "Накопительная таблица");

        }
    }
}
