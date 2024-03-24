using Data;
using Files.Abstracts;
using Files.Repositorys;
using Files.Tables;
using Logging;
using Microsoft.Office.Interop.Excel;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Files
{
    enum Crossing{
        BIG,SMALL
    };
    public class SurveyReportsFileHandler : FileHandler
    {
        public bool ErrorFiles {  get; set; }
        private string RootFolder { get; set; }
        private DataAnalyzer DataAnalyzer { get; set; }
        private MoscowTable MoscowTable { get; set; }
        private PiktsTable PiktsTable { get; set; }
        private Repository<Queue<UnderwaterCrossing>, UnderwaterCrossing> Repository { get; set; }
        private Repository<Queue<PIKTS>, PIKTS> PiktsData { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public SurveyReportsFileHandler(DataAnalyzer dataAnalyzer, MoscowTable moscowTable,
            string rootFolder, PiktsTable piktsTable, IProgress<object[]> progress)
        {
            ArgumentNullException.ThrowIfNull(dataAnalyzer);
            ArgumentNullException.ThrowIfNull(moscowTable);
            ArgumentNullException.ThrowIfNull(rootFolder);
            ArgumentNullException.ThrowIfNull(piktsTable);

            DataAnalyzer = dataAnalyzer;
            MoscowTable = moscowTable;
            RootFolder = rootFolder;
            PiktsTable = piktsTable;
            Progress = progress;

            if (!Directory.Exists(RootFolder))
                throw new ArgumentException("Папка с приложениями А не найдена");

            Repository = new SurveyReportsRepository();
            PiktsData = new PIKTSRepository();
        }
        public override void HandleFiles()
        {
            try
            {
                ReadAnnexes();
                FillMoscowTable();

                ExcelApp.Run.Quit();
            }
            finally
            {
                if (ExcelApp.Run != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ExcelApp.Run);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }

        }
        /// <summary>
        /// Чтение файлов Приложение А
        /// </summary>
        private void ReadAnnexes()
        {
            string[] Files = Directory.GetFiles(RootFolder);
            int indexProgress = 1;

            foreach (string file in Files)
            {
                Progress.Report(new object[] { Files.Length, indexProgress++, "Чтение файлов Приложений А...", 1 });
                try
                {
                    UnderwaterCrossing underwaterCrossing = new();

                    switch (CheckTypeOfCrossing(file))
                    {
                        case Crossing.BIG:
                            ReadAsBig(file, underwaterCrossing);
                            break;
                        case Crossing.SMALL:
                            ReadAsSmall(file, underwaterCrossing);
                            break;
                        default:
                            break;
                    }
                }
                catch (IOException ex)
                {
                    Loger.Log(file, ex.Message);
                }
                catch (ArgumentException ex)
                {
                    Loger.Log(file, ex.Message);
                }
                catch (Exception ex)
                {
                    Loger.Log(file, ex.Message);
                }
            }
        }
        /// <summary>
        /// Чтение файла ППМТ (большой водоток)
        /// </summary>
        /// <param name="pathFile"></param>
        private void ReadAsBig(string pathFile, UnderwaterCrossing underwaterCrossing)
        {
            // Создание текущей рабочей книги
            Excel.Workbook curWorkbook = ExcelApp.Run.Workbooks.Open(pathFile);

            // Чтение всех листов (вкладок) рабочей книги
            List<Excel.Worksheet> worksheets = ReadAllSheets(curWorkbook);

            // Получение данных ячеек листа 4.Тех.хар-ка МТ 
            Excel.Worksheet curWorksheet = FindSheet(worksheets, "хар-ка", pathFile);
            Excel.Range cell = FindCell(curWorksheet, "ID", pathFile);
            underwaterCrossing.Id = DataAnalyzer.
                GetIdValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), pathFile);
            if (underwaterCrossing.Id == null) 
            {
                Loger.Log(pathFile, $"Данные из Приложения А. В московской таблице (лист ППМН ТН) " +
                    $"отсутствует переход с ID - {underwaterCrossing.Id}");
                return;
            } 
            // Получение данных ячеек листа 1.Общие сведения
            curWorksheet = FindSheet(worksheets, "Общие сведения", pathFile);
            cell = FindCell(curWorksheet, "Дата обследования", pathFile);
            underwaterCrossing.DateInspection = DataAnalyzer.
                GetDateValue(curWorksheet.Cells[cell.Row + 2, cell.Column].Value.ToString(), pathFile);
            underwaterCrossing.TypeOfSurvey = "обследование";

            // Получение данных ячеек листа 9.ФГХ
            curWorksheet = FindSheet(worksheets, "ФГХ", pathFile);
            cell = FindCell(curWorksheet, "Широта, левый берег", pathFile);
            underwaterCrossing.Coordinates!.
                LeftCoastLatitude = DataAnalyzer.
                ConvertGeoDegreeToDouble(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), pathFile, cell);
            cell = FindCell(curWorksheet, "Долгота, левый берег", pathFile);
            underwaterCrossing.Coordinates.
                LeftCoastLongitude = DataAnalyzer.
                ConvertGeoDegreeToDouble(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), pathFile, cell);
            cell = FindCell(curWorksheet, "Широта, правый берег", pathFile);
            underwaterCrossing.Coordinates.
                RightCoastLatitude = DataAnalyzer.
                ConvertGeoDegreeToDouble(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), pathFile, cell);
            cell = FindCell(curWorksheet, "Долгота, правый берег", pathFile);
            underwaterCrossing.Coordinates.
                RightCoastLongitude = DataAnalyzer.
                ConvertGeoDegreeToDouble(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), pathFile, cell);

            // Получение данных ячеек листа 12.1% и 10%, Нмеж.
            curWorksheet = FindSheet(worksheets, "Нмеж", pathFile);
            cell = FindCell(curWorksheet, "Расход воды", pathFile);
            underwaterCrossing.WaterRate!.OnePercentLevel = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            underwaterCrossing.WaterRate.TenPercentLevel = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 3].Value.ToString(), cell, pathFile);
            underwaterCrossing.WaterRate.AverageLevel = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 4].Value.ToString(), cell, pathFile);
            cell = FindCell(curWorksheet, "Максимальная скорость течения воды в русле", pathFile);
            underwaterCrossing.MaxSpeeds!.MaxSpeedOnePercent = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            underwaterCrossing.MaxSpeeds.MaxSpeedTenPercent = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 3].Value.ToString(), cell, pathFile);
            underwaterCrossing.MaxSpeeds.MaxSpeedAverage = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 4].Value.ToString(), cell, pathFile);

            // Получение данных ячеек листа 15.РП и ППРР
            curWorksheet = FindSheet(worksheets, "ППРР", pathFile);
            cell = FindCell(curWorksheet, "Залегание МТ относительно профиля", pathFile);
            underwaterCrossing.PositionMT =
                curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString();

            // Получение данных ячеек листа 16.Анализ РП и ПВП
            curWorksheet = FindSheet(worksheets, "Анализ", pathFile);
            cell = FindCell(curWorksheet, "заложения", pathFile);
            underwaterCrossing.DeviationsPVP!.NGZRiverbed = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            underwaterCrossing.DeviationsPVP.NGZFloodplain = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 2].Value.ToString(), cell, pathFile);
            cell = FindCell(curWorksheet, "Длина оголения", pathFile);
            underwaterCrossing.DeviationsPVP.DenudationRiverbed = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            underwaterCrossing.DeviationsPVP.DenudationFloodplain = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 2].Value.ToString(), cell, pathFile);
            cell = FindCell(curWorksheet, "Длина провиса", pathFile);
            underwaterCrossing.DeviationsPVP.SagRiverbed = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            underwaterCrossing.DeviationsPVP.SagFloodplain = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 2].Value.ToString(), cell, pathFile);
            cell = FindCell(curWorksheet, "Скорость", pathFile);
            underwaterCrossing.RivebedProcesses!.SpeedOffsetRiverbed = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            cell = FindCell(curWorksheet, "Амплитуда", pathFile);
            underwaterCrossing.RivebedProcesses.AmplitudeRiverbed = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            cell = FindCell(curWorksheet, "Высота", pathFile);
            underwaterCrossing.RivebedProcesses.HeightMicroforms = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);

            // Определение характера русловых процессов
            underwaterCrossing.Character = DataAnalyzer.
                GetCharacterOfProcess(underwaterCrossing.RivebedProcesses, pathFile);

            // Добавление в репозиторий
            Repository.Put(underwaterCrossing);

            // Закрытие текущей книги
            curWorkbook.Close(false);
        }
        /// <summary>
        /// Чтение файла МВ (малый водоток)
        /// </summary>
        /// <param name="pathFile"></param>
        private void ReadAsSmall(string pathFile, UnderwaterCrossing underwaterCrossing)
        {
            // Создание текущей рабочей книги
            Excel.Workbook curWorkbook = ExcelApp.Run.Workbooks.Open(pathFile);

            // Чтение всех листов (вкладок) рабочей книги
            List<Excel.Worksheet> worksheets = ReadAllSheets(curWorkbook);

            // Получение данных ячеек листа 4.Тех.хар-ка МТ
            Excel.Worksheet curWorksheet = FindSheet(worksheets, "хар-ка", pathFile);
            Excel.Range cell = FindCell(curWorksheet, "ID", pathFile);
            underwaterCrossing.Id = DataAnalyzer.
                GetIdValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), pathFile);
            if (underwaterCrossing.Id == null)
            {
                Loger.Log(pathFile, $"Данные из Приложения А. В московской таблице (лист ППМН ТН) " +
                    $"отсутствует переход с ID - {underwaterCrossing.Id}");
                return;
            }

            // Получение данных ячеек листа 1.Общие сведения
            curWorksheet = FindSheet(worksheets, "Общие сведения", pathFile);
            cell = FindCell(curWorksheet, "Дата обследования", pathFile);
            underwaterCrossing.DateInspection = DataAnalyzer.
                GetDateValue(curWorksheet.Cells[cell.Row + 2, cell.Column].Value.ToString(), pathFile);
            underwaterCrossing.TypeOfSurvey = "обследование";

            // Получение данных ячеек листа 9.ФГХ
            curWorksheet = FindSheet(worksheets, "ФГХ", pathFile);
            cell = FindCell(curWorksheet, "Широта", pathFile);
            underwaterCrossing.Coordinates!.LeftCoastLatitude = DataAnalyzer.
                ConvertGeoDegreeToDouble(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), pathFile, cell);
            cell = FindCell(curWorksheet, "Долгота", pathFile);
            underwaterCrossing.Coordinates.RightCoastLongitude = DataAnalyzer.
                ConvertGeoDegreeToDouble(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), pathFile, cell);

            // Получение данных ячеек листа 13.Анализ ПВП
            curWorksheet = FindSheet(worksheets, "Анализ", pathFile);
            cell = FindCell(curWorksheet, "заложения", pathFile);
            underwaterCrossing.DeviationsPVP!.NGZRiverbed = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            underwaterCrossing.DeviationsPVP.NGZFloodplain = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 2].Value.ToString(), cell, pathFile);
            cell = FindCell(curWorksheet, "Длина оголения", pathFile);
            underwaterCrossing.DeviationsPVP.DenudationRiverbed = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            underwaterCrossing.DeviationsPVP.DenudationFloodplain = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 2].Value.ToString(), cell, pathFile);
            cell = FindCell(curWorksheet, "Длина провиса", pathFile);
            underwaterCrossing.DeviationsPVP.SagRiverbed = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 1].Value.ToString(), cell, pathFile);
            underwaterCrossing.DeviationsPVP.SagFloodplain = DataAnalyzer.
                GetFloatValue(curWorksheet.Cells[cell.Row, cell.Column + 2].Value.ToString(), cell, pathFile);

            // Получение данных ячеек листа 6.Ремонты
            curWorksheet = FindSheet(worksheets, "Ремонт", pathFile);
            cell = FindCell(curWorksheet, "выполненных работ", pathFile);
            underwaterCrossing.RepairInfo = DataAnalyzer.GetRepairInfo(curWorksheet, cell);

            // Получение данных ячеек листа 15.1 ПВП недозаглубления
            curWorksheet = FindSheet(worksheets, "недозаглубления", pathFile);
            DataAnalyzer.GetDeviationsRivebed(curWorksheet, "недозаглубления", underwaterCrossing);

            // Получение данных ячеек листа 15.2 ПВП оголения
            curWorksheet = FindSheet(worksheets, "оголения", pathFile);
            DataAnalyzer.GetDeviationsRivebed(curWorksheet, "оголения", underwaterCrossing);

            // Получение данных ячеек листа 15.3 ПВП провисы
            curWorksheet = FindSheet(worksheets, "провисы", pathFile);
            DataAnalyzer.GetDeviationsRivebed(curWorksheet, "провисы", underwaterCrossing);

            //Получение значения положения МТ по отношению к ППР
            underwaterCrossing.PositionMT = DataAnalyzer.
                GetPositionPipelineFor(underwaterCrossing.DeviationsPVP);

            // Добавление в репозиторий
            Repository.Put(underwaterCrossing);

            // Закрытие текущей книги
            curWorkbook.Close(false);
        }
        /// <summary>
        /// Получение всех вкладок Excel файла
        /// </summary>
        private List<Excel.Worksheet> ReadAllSheets(Excel.Workbook curWorkbook)
        {
            List<Excel.Worksheet> worksheets = new();
            foreach (Worksheet worksheet in curWorkbook.Worksheets)
                worksheets.Add(worksheet);
            return worksheets;
        }
        /// <summary>
        /// Поиск листа в книге
        /// </summary>
        /// <param name="value">Имя листа Excel книги</param>
        /// <exception cref="ArgumentNullException">Лист не найден</exception>
        private Excel.Worksheet FindSheet(List<Excel.Worksheet> worksheets, string value, string filePath)
        {
            Excel.Worksheet? worksheet = worksheets.Find(sheet => sheet.Name.
                    Contains(value, StringComparison.OrdinalIgnoreCase));
            if (worksheet != null)
                return worksheet;
            else
            {
                Loger.Log(filePath, $"Имя листа в книге не найдено: {value}");
                throw new ArgumentNullException();
            }
        }
        /// <summary>
        /// Поиск ячейки на листе
        /// </summary>
        /// <param name="value">Значение ячейки</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">Ячейка не найдена</exception>
        private Excel.Range FindCell(Excel.Worksheet curWorksheet, string value, string filePath)
        {
            Excel.Range cell = curWorksheet.Cells.
                Find(value, Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlPart);
            if (cell != null)
                return cell;
            else
            {
                Loger.Log(filePath, $"Ячейка с именем: {value} не найдена");
                throw new ArgumentNullException();
            }
        }
        /// <summary>
        /// Проверка типа перехода - ППМТ или МВ
        /// </summary>
        private Crossing CheckTypeOfCrossing(string file)
        {
            Excel.Workbook curWorkbook = ExcelApp.Run.Workbooks.Open(file);

            List<Excel.Worksheet> worksheets = ReadAllSheets(curWorkbook);

            Excel.Worksheet curWorksheet = FindSheet(worksheets, "хар-ка", file);

            Excel.Range cell = curWorksheet.Cells.
                Find("Наличие судоходности", Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlPart);

            curWorkbook.Close(false);

            if (cell != null)
                return Crossing.BIG;
            else
                return Crossing.SMALL;

        }
        /// <summary>
        /// Получение данных из ПИКТС таблицы (отчет 31.ПАО)
        /// </summary>
        private void GetInfoFromPiktsTable()
        {
            int indexProgress = 1;
            System.Data.DataTable piktsTable = DataAnalyzer.ArrayToDataTable(PiktsTable.TableData);

            IEnumerable<string?> IDs = piktsTable.
                AsEnumerable().
                Select(coll => coll?.Field<string>(PiktsTable.PiktsTableColumns["ID ППМТ"][0] - 1)).
                Distinct().
                Where(value => value?.Length == 6);

            foreach (string? id in IDs)
            {
                PIKTS piktsParams = new PIKTS();

                if (id == null) continue;

                IEnumerable<DataRow> dataRows = piktsTable.
                Select().
                Where(row => row.Field<string>(PiktsTable.PiktsTableColumns["ID ППМТ"][0] - 1) == id);

                if (!dataRows.Any()) continue;

                System.Data.DataTable dataTable = dataRows.CopyToDataTable();

                piktsParams.ID = id;

                piktsParams.ConditionOfPipeline =
                    dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Состояние нитки"][0] - 1);

                piktsParams.SteelGrade =
                    dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Характеристика труб руслового участка"][0] - 1);

                piktsParams.DateLastVTD =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата прогона"][0] - 1));

                Func<string, string, StringComparison, string> countDefects = delegate (string year, string location, StringComparison comparisonType)
                {
                    int quantityDefects;
                    try
                    {
                        quantityDefects = dataTable.Select().
                        Where(row => row.ItemArray[PiktsTable.PiktsTableColumns["Предельная дата эксплуатации трубной секции с дефектом"][0] - 1] != DBNull.Value
                        &&
                        row.ItemArray[PiktsTable.PiktsTableColumns["Признак ремонта"][0] - 1] != DBNull.Value
                        &&
                        row.ItemArray[PiktsTable.PiktsTableColumns["Местоположение дефекта (русло/пойма)"][0] - 1] != DBNull.Value).
                        Count(row =>
                        row.Field<string>(PiktsTable.PiktsTableColumns["Предельная дата эксплуатации трубной секции с дефектом"][0] - 1).
                        Contains(year)
                        &&
                        row.Field<string>(PiktsTable.PiktsTableColumns["Признак ремонта"][0] - 1).Equals("Без ремонта", comparisonType)
                        &&
                        row.Field<string>(PiktsTable.PiktsTableColumns["Местоположение дефекта (русло/пойма)"][0] - 1).Equals(location, comparisonType));

                        return quantityDefects == 0 ? "" : quantityDefects.ToString();
                    }
                    catch (NullReferenceException ex)
                    {
                        Loger.Log(ex.Message);
                    }
                    catch (Exception ex)
                    {
                        Loger.Log(ex.Message);
                    }

                    return "";
                };

                piktsParams.DefectsPoima[DateTime.Now.Year.ToString()] =
                    countDefects(DateTime.Now.Year.ToString(), "Пойма", StringComparison.OrdinalIgnoreCase);

                piktsParams.DefectsPoima[(DateTime.Now.Year + 1).ToString()] =
                    countDefects((DateTime.Now.Year + 1).ToString(), "Пойма", StringComparison.OrdinalIgnoreCase);

                piktsParams.DefectsPoima[(DateTime.Now.Year + 2).ToString()] =
                    countDefects((DateTime.Now.Year + 2).ToString(), "Пойма", StringComparison.OrdinalIgnoreCase);

                piktsParams.DefectsRuslo[DateTime.Now.Year.ToString()] =
                    countDefects(DateTime.Now.Year.ToString(), "Русло", StringComparison.OrdinalIgnoreCase);

                piktsParams.DefectsRuslo[(DateTime.Now.Year + 1).ToString()] =
                    countDefects((DateTime.Now.Year + 1).ToString(), "Русло", StringComparison.OrdinalIgnoreCase);

                piktsParams.DefectsRuslo[(DateTime.Now.Year + 2).ToString()] =
                    countDefects((DateTime.Now.Year + 2).ToString(), "Русло", StringComparison.OrdinalIgnoreCase);

                piktsParams.SafePeriod =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации"][0] - 1));

                piktsParams.OTSNumber =
                    dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Номер заключения по оценке технического состояния"][0] - 1);

                piktsParams.DateReport =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата выдачи заключения по оценке технического состояния"][0] - 1));

                piktsParams.DateVTD =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][0] - 1));

                piktsParams.DateDefect =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][1] - 1));

                piktsParams.DateCDS =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][2] - 1));

                piktsParams.DateJumpers =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][3] - 1));

                piktsParams.DateLimited =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][4] - 1));

                piktsParams.DateVRK =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][5] - 1));

                piktsParams.DateUZA =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][6] - 1));

                piktsParams.DateWeldedElement =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][7] - 1));

                piktsParams.DateConnectedDetails =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][8] - 1));

                piktsParams.DateKPSOD =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][9] - 1));

                piktsParams.DateDrainageContainers =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][9] - 1));

                piktsParams.DatePVP =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][10] - 1));

                piktsParams.DateCorrosion =
                    DataAnalyzer.ExtractDate(dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Дата окончания срока безопасной эксплуатации по параметрам"][11] - 1));

                piktsParams.Organization =
                    dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Организация, проводившая экспертную оценку"][0] - 1);

                piktsParams.Events =
                    dataTable.Rows[0].Field<string>(PiktsTable.PiktsTableColumns["Рекомендации по приведению ПМТ в нормативное состояние"][0] - 1) != null ? "Да" : null;

                PiktsData.Put(piktsParams);

                Progress.Report(new object[] { IDs.Count(), indexProgress++, "Чтение таблицы ПИКТС...", 1 });
            }
        }
        /// <summary>
        /// Заполнение московской таблицы
        /// </summary>
        private void FillMoscowTable()
        {
            int quantAnnexes = Repository.Storage.Count;
            int indexProgress = 1;
            UnderwaterCrossing underwaterCrossing;
            Excel.Range? ID;

            while (Repository.Storage.Count > 0)
            {
                Progress.Report(new object[] { quantAnnexes, indexProgress++, "Заполнение московской таблицы...", 1 });
                try
                {
                    underwaterCrossing = Repository.Get();
                    ID = SearchID(underwaterCrossing);

                    switch (ID?.Worksheet.Name)
                    {
                        case "ППМН ТН":
                            FillAsBig(underwaterCrossing, ID);
                            break;
                        case "МВ ТН":
                            FillAsSmall(underwaterCrossing, ID);
                            break;
                        default:
                            Loger.Log($"{underwaterCrossing.Id} в московской таблице не найден");
                            break;
                    }
                }
                catch (IOException ex)
                {
                    Loger.Log(MoscowTable.Path, ex.Message);
                }
                catch (Exception ex)
                {
                    Loger.Log(MoscowTable.Path, ex.Message);
                }
            }
            if(ErrorFiles == false)
            {
                try
                {
                    GetInfoFromPiktsTable();
                    TransferPiktsData();
                }
                catch (IOException ex)
                {
                    Loger.Log(MoscowTable.Path, ex.Message);
                }
                catch (Exception ex)
                {
                    Loger.Log(MoscowTable.Path, ex.Message);
                }
            }
        }
        /// <summary>
        /// Поиск ID перехода в московской таблице
        /// </summary>
        private Excel.Range? SearchID(UnderwaterCrossing underwaterCrossing)
        {
            Excel.Workbook curWorkbook = ExcelApp.Run.Workbooks.Open(MoscowTable.Path);

            Excel.Range? IDcell;

            IDcell = curWorkbook.Worksheets["ППМН ТН"].Range
                [
                curWorkbook.Worksheets["ППМН ТН"].Cells[1, MoscowTable.BigTableColumns["ID"][0]],
                curWorkbook.Worksheets["ППМН ТН"].Cells[curWorkbook.Worksheets["ППМН ТН"].UsedRange.Rows.Count, MoscowTable.BigTableColumns["ID"][0]]
                ].Cells.
                Find(underwaterCrossing.Id, Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole);


            if (IDcell != null)
                return IDcell;
            else
            {
                IDcell = curWorkbook.Worksheets["МВ ТН"].Range
                [
                curWorkbook.Worksheets["МВ ТН"].Cells[1, MoscowTable.BigTableColumns["ID"][0]],
                curWorkbook.Worksheets["МВ ТН"].Cells[curWorkbook.Worksheets["МВ ТН"].UsedRange.Rows.Count, MoscowTable.SmallTableColumns["ID"][0]]
                ].Cells.
                Find(underwaterCrossing.Id, Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole);
            }

            return IDcell;
        }
        /// <summary>
        /// Заполнение листа ППМТ - большой водоток
        /// </summary>
        private void FillAsBig(UnderwaterCrossing underwaterCrossing, Excel.Range ID)
        {
            Excel.Worksheet curWorksheet = ExcelApp.Run.Worksheets["ППМН ТН"];

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Дата последнего обследования"][0]]
                .Value = underwaterCrossing.DateInspection.ToString();

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Вид обслед."][0]]
                .Value = underwaterCrossing.TypeOfSurvey ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Положение МТ по отношению к ППРР (выше/ниже)"][0]]
                .Value = underwaterCrossing.PositionMT ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Разделение на пойму и русло"][0]]
                .Value = underwaterCrossing.DeviationsPVP?.NGZRiverbed ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Разделение на пойму и русло"][1]]
                .Value = underwaterCrossing.DeviationsPVP.DenudationRiverbed ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Разделение на пойму и русло"][2]]
                .Value = underwaterCrossing.DeviationsPVP.SagRiverbed ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Разделение на пойму и русло"][3]]
                .Value = underwaterCrossing.DeviationsPVP.NGZFloodplain ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Разделение на пойму и русло"][4]]
                .Value = underwaterCrossing.DeviationsPVP.DenudationFloodplain ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Разделение на пойму и русло"][5]]
                .Value = underwaterCrossing.DeviationsPVP.SagFloodplain ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Русловые процессы"][0]]
                .Value = underwaterCrossing.RivebedProcesses.SpeedOffsetRiverbed ?? 0; 

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Русловые процессы"][1]]
                .Value = underwaterCrossing.RivebedProcesses.AmplitudeRiverbed ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Русловые процессы"][2]]
                .Value = underwaterCrossing.RivebedProcesses.HeightMicroforms ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Русловые процессы"][3]]
                .Value = underwaterCrossing.Character ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Левый берег"][0]].numberformatlocal = "@";
            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Левый берег"][0]]
                .Value = underwaterCrossing.Coordinates.LeftCoastLatitude ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Левый берег"][1]].numberformatlocal = "@";
            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Левый берег"][1]]
                .Value = underwaterCrossing.Coordinates.LeftCoastLongitude ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Правый берег"][0]].numberformatlocal = "@";
            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Правый берег"][0]]
                .Value = underwaterCrossing.Coordinates.RightCoastLatitude ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Правый берег"][1]].numberformatlocal = "@";
            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Правый берег"][1]]
                .Value = underwaterCrossing.Coordinates.RightCoastLongitude ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Расход воды, м³/с"][0]]
                .Value = underwaterCrossing.WaterRate.OnePercentLevel ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Расход воды, м³/с"][1]]
                .Value = underwaterCrossing.WaterRate.TenPercentLevel ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Расход воды, м³/с"][2]]
                .Value = underwaterCrossing.WaterRate.AverageLevel ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Максимальная скорость течения воды в русле, м/с"][0]]
                .Value = underwaterCrossing.MaxSpeeds.MaxSpeedOnePercent ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Максимальная скорость течения воды в русле, м/с"][1]]
                .Value = underwaterCrossing.MaxSpeeds.MaxSpeedTenPercent ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Максимальная скорость течения воды в русле, м/с"][2]]
                .Value = underwaterCrossing.MaxSpeeds.MaxSpeedAverage ?? 0;

            ExcelApp.Run.Workbooks[1].Save();
        }
        /// <summary>
        /// Заполнение листа МВ - малый водоток
        /// </summary>
        private void FillAsSmall(UnderwaterCrossing underwaterCrossing, Excel.Range ID)
        {
            Excel.Worksheet curWorksheet = ExcelApp.Run.Worksheets["МВ ТН"];

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Дата последнего обследования"][0]]
                .Value = underwaterCrossing?.DateInspection?.ToString() ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Вид обследования"][0]]
                .Value = underwaterCrossing.TypeOfSurvey ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Положение МТ по отношению к ППРР (выше/ниже)"][0]]
                .Value = underwaterCrossing.PositionMT ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Разделение на пойму и русло"][0]]
                .Value = underwaterCrossing.DeviationsPVP.NGZRiverbed ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Разделение на пойму и русло"][1]]
                .Value = underwaterCrossing.DeviationsPVP.DenudationRiverbed ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Разделение на пойму и русло"][2]]
                .Value = underwaterCrossing.DeviationsPVP.SagRiverbed ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Разделение на пойму и русло"][3]]
                .Value = underwaterCrossing.DeviationsPVP.NGZFloodplain ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Разделение на пойму и русло"][4]]
                .Value = underwaterCrossing.DeviationsPVP.DenudationFloodplain ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Разделение на пойму и русло"][5]]
                .Value = underwaterCrossing.DeviationsPVP.SagFloodplain ?? 0;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["информация о ремонтах выявленных отклонений"][0]]
                .Value = underwaterCrossing.RepairInfo ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Широта на момент обследования "][0]].numberformatlocal = "@";
            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Широта на момент обследования "][0]]
                .Value = underwaterCrossing.Coordinates.LeftCoastLatitude ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Долгота на момент обследования "][0]].numberformatlocal = "@";
            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Долгота на момент обследования "][0]]
                .Value = underwaterCrossing.Coordinates.RightCoastLongitude ?? "";

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Наличие отклонений ПВП в русле"][0]]
                .Value = underwaterCrossing.DeviationsRivebed.LengthNGZ;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Наличие отклонений ПВП в русле"][1]]
                .Value = underwaterCrossing.DeviationsRivebed.MinThicknessProtectingLayer;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Наличие отклонений ПВП в русле"][2]]
                .Value = underwaterCrossing.DeviationsRivebed.LengthDenudation;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Наличие отклонений ПВП в русле"][3]]
                .Value = underwaterCrossing.DeviationsRivebed.MaxDepthDenudation;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Наличие отклонений ПВП в русле"][4]]
                .Value = underwaterCrossing.DeviationsRivebed.LengthSag;

            curWorksheet.Cells[ID.Row, MoscowTable.SmallTableColumns["Наличие отклонений ПВП в русле"][5]]
                .Value = underwaterCrossing.DeviationsRivebed.MaxLengthSinglePart;

            ExcelApp.Run.Workbooks[1].Save();
        }
        /// <summary>
        /// Перенос данных из ПИКТС (отчет 31.ПАО) в московскую таблицу
        /// </summary>
        private void TransferPiktsData()
        {
            int quantParams = PiktsData.Storage.Count;
            int indexProgress = 1;
            Excel.Worksheet curWorksheet = ExcelApp.Run.Worksheets["ППМН ТН"];

            while (PiktsData.Storage.Count > 0)
            {
                Progress.Report(new object[] { quantParams, indexProgress++, "Перенос данных из ПИКТС таблицы...", 1 });

                PIKTS piktsParams = PiktsData.Get();

                Excel.Range? ID = curWorksheet.Range
                [
                curWorksheet.Cells[1, MoscowTable.BigTableColumns["ID"][0]],
                curWorksheet.Cells[curWorksheet.UsedRange.Rows.Count, MoscowTable.BigTableColumns["ID"][0]]
                ].Cells.
                Find(piktsParams.ID, Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole);

                if (ID == null)
                {
                    Loger.Log($"Данные из ПИКТС. В московской таблице (лист ППМН ТН) отсутствует переход с ID - {piktsParams.ID}");
                    continue;
                }

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["состояние (в работе/отключена)"][0]]
                .Value = piktsParams.ConditionOfPipeline ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Марка стали"][0]]
                .Value = piktsParams.SteelGrade ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Дата  последнего ВТД"][0]]
                .Value = piktsParams.DateLastVTD ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Наличие дефектов"][0]]
                    .Value = piktsParams.DefectsPoima[DateTime.Now.Year.ToString()] ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Наличие дефектов"][1]]
                    .Value = piktsParams.DefectsRuslo[DateTime.Now.Year.ToString()] ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Наличие дефектов"][2]]
                    .Value = piktsParams.DefectsPoima[(DateTime.Now.Year + 1).ToString()] ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Наличие дефектов"][3]]
                    .Value = piktsParams.DefectsRuslo[(DateTime.Now.Year + 1).ToString()] ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Наличие дефектов"][4]]
                    .Value = piktsParams.DefectsPoima[(DateTime.Now.Year + 2).ToString()] ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Наличие дефектов"][5]]
                    .Value = piktsParams.DefectsRuslo[(DateTime.Now.Year + 2).ToString()] ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Срок безопасной эксплуатации по отчету ОТС, дата (ЧЧ.ММ.ГГ)"][0]]
                    .Value = piktsParams.SafePeriod ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Количество ОТС"][1]]
                    .Value = piktsParams.DateReport ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Номер ОТС"][0]]
                    .Value = piktsParams.OTSNumber ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][0]]
                    .Value = piktsParams.DateVTD ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][1]]
                    .Value = piktsParams.DateDefect ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][2]]
                    .Value = piktsParams.DateCDS ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][3]]
                    .Value = piktsParams.DateJumpers ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][4]]
                    .Value = piktsParams.DateLimited ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][5]]
                    .Value = piktsParams.DateVRK ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][6]]
                    .Value = piktsParams.DateUZA ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][7]]
                    .Value = piktsParams.DateWeldedElement ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][8]]
                    .Value = piktsParams.DateConnectedDetails ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][9]]
                    .Value = piktsParams.DateKPSOD ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][10]]
                    .Value = piktsParams.DateDrainageContainers ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][11]]
                    .Value = piktsParams.DatePVP ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Cрок безопасной эксплуатации по параметрам, шт."][12]]
                    .Value = piktsParams.DateCorrosion ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Организация, выдавшая заключение о сроке безопасной эксплуатации"][0]]
                    .Value = piktsParams.Organization ?? "";

                curWorksheet.Cells[ID.Row, MoscowTable.BigTableColumns["Наличие мероприятий по приведению в норм сост. да/нет"][0]]
                    .Value = piktsParams.Events ?? "";

                ExcelApp.Run.Workbooks[1].Save();
            }
        }
    }
}
