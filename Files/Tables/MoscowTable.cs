namespace Files.Tables
{
    public class MoscowTable
    {
        internal string Path { get; set; }

        internal Dictionary<string, List<int>> BigTableColumns = new Dictionary<string, List<int>>()
        {
            {"состояние (в работе/отключена)", new() },
            {"Дата последнего обследования", new() },
            {"Вид обслед.",new() },
            {"Положение МТ по отношению к ППРР (выше/ниже)",new() },
            {"Разделение на пойму и русло",new() },
            {"Русловые процессы",new() },
            {"Марка стали",new() },
            {"Дата  последнего ВТД",new() },
            {"Наличие дефектов",new() },
            {"Срок безопасной эксплуатации по отчету ОТС, дата (ЧЧ.ММ.ГГ)",new() },
            {"Количество ОТС",new() },
            {"Номер ОТС",new() },
            {"Cрок безопасной эксплуатации по параметрам, шт.", new() },
            {"Организация, выдавшая заключение о сроке безопасной эксплуатации",new() },
            {"Наличие мероприятий по приведению в норм сост. да/нет",new() },
            {"Левый берег",new() },
            {"Правый берег",new() },
            {"ID",new() },
            {"Расход воды, м³/с",new() },
            {"Максимальная скорость течения воды в русле, м/с",new() }
        };

        internal Dictionary<string, List<int>> SmallTableColumns = new Dictionary<string, List<int>>()
        {
            {"Дата последнего обследования",new() },
            {"Вид обследования",new() },
            {"Положение МТ по отношению к ППРР (выше/ниже)",new() },
            {"Разделение на пойму и русло",new() },
            {"информация о ремонтах выявленных отклонений",new() },
            {"Широта на момент обследования ",new() },
            {"Долгота на момент обследования ",new() },
            {"ID",new() },
            {"Наличие отклонений ПВП в русле",new() },
        };
        public MoscowTable(DataAnalyzer dataAnalyzer, string path)
        {
            ArgumentNullException.ThrowIfNull(dataAnalyzer);
            ArgumentNullException.ThrowIfNull(path, "Выберите файл с московской таблицей");

            Path = path;
            if (!File.Exists(Path)) throw new ArgumentException("файл Московской таблицы не существует");

            dataAnalyzer.GetHeaderColumns(Path, BigTableColumns, "ППМН ТН");
            dataAnalyzer.GetHeaderColumns(Path, SmallTableColumns, "МВ ТН");
        }
    }
}
