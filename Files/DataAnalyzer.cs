using Data;
using Data.CrossingParts;
using Logging;
using System.Data;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Files
{
    public class DataAnalyzer
    {
        /// <summary>
        /// Определение характера русловых процессов
        /// </summary>
        internal string GetCharacterOfProcess(RivebedProcesses processes, string pathFile)        {
            bool isNullProp = typeof(RivebedProcesses).GetProperties().
                Any(prop => prop.GetValue(processes) == null);

            if (isNullProp)
            {
                Loger.Log(pathFile, "Проверьте значения русловых процессов");
            }

            if (processes.SpeedOffsetRiverbed > 2 || 
                processes.AmplitudeRiverbed > 1 || 
                processes.HeightMicroforms > 1.5)
                return "интенсивный";
            else if (processes.SpeedOffsetRiverbed < 0.5 &&
                processes.AmplitudeRiverbed <= 1 &&
                processes.HeightMicroforms < 0.5)
                return "стабильный";
            else
                return "умеренный";
        }
        /// <summary>
        /// Получение корректного ID подводного перехода
        /// </summary>
        internal string? GetIdValue(string value, string pathFile)
        {
            if (value == null) return null;

            int id;
            int.TryParse(string.
                Join("", value.
                Where(symbol => char.IsDigit(symbol))), out id);

            if (id == 0)
            {
                Loger.Log(pathFile, "Проверьте ID перехода (лист 4)");
                return null;
            }
            else
                return id.ToString();
        }
        /// <summary>
        /// Получение корректной даты обследования
        /// </summary>
        internal DateOnly? GetDateValue(string value, string pathFile)
        {
            DateTime date;
            if (DateTime.TryParse(value, out date) == false)
            {
                Loger.Log(pathFile, "Проверьте дату обследования (лист 1)");
                return null;
            }
            else
                return DateOnly.FromDateTime(date);
        }
        /// <summary>
        /// Извлечение даты из строки
        /// </summary>
        /// <param name="value">строка с датой и временем</param>
        /// <returns></returns>
        internal string? ExtractDate(string? value)
        {
            if (value == null) return null;

            Regex onlyDate = new Regex(@"\b\d{2}\.\d{2}.\d{4}\b");
            
            if (onlyDate.IsMatch(value))
                return onlyDate.Match(value).ToString();
            else 
                return value;
        }
        /// <summary>
        /// Получение корректного значения числа с плавающей точкой
        /// </summary>
        internal float? GetFloatValue(string? value, Excel.Range cell, string pathFile)
        {
            if (value == null) return null;

            float number;

            if (float.TryParse(value, out number) == false)
                try
                {
                    value = value.Replace("*", "");
                    return float.Parse(Regex.Replace(value, @"\.+", ","));
                }
                catch (FormatException)
                {
                    if(cell.Worksheet.Name != "12.1% и 10%, Нмеж.")
                        Loger.Log(pathFile, cell, $"Проверьте значение: {value}");
                    
                    return null;
                }
            else
                return number;
        }
        /// <summary>
        /// Получение корректного значения числа с плавающей точкой
        /// </summary>
        internal float? GetFloatValue(string? value)
        {
            if (value == null) return null;

            float number;

            if (float.TryParse(value, out number) == false)
                try
                {
                    return float.Parse(Regex.Replace(value, @"\.+", ","));
                }
                catch (FormatException)
                {
                    return null;
                }
            else
                return number;
        }
        /// <summary>
        /// Получение сведений о ремонтах на переходе
        /// </summary>
        /// <param name="curWorksheet"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        internal string GetRepairInfo(Excel.Worksheet curWorksheet, Excel.Range cell)
        {
            string value = "";

            object[,] data = curWorksheet.
                Range[curWorksheet.Cells[cell.Row + 1, cell.Column], 
                      curWorksheet.Cells[curWorksheet.UsedRange.Rows.Count, cell.Column]
                      ].Cells.Value;

            foreach (object item in data)
            {
                if (item != null && item is string)
                    value = value + (string)item  + ". ";
            }
            return value;
        }
        /// <summary>
        /// Получение критич. данных по руслу реки
        /// </summary>
        /// <returns></returns>
        internal void GetDeviationsRivebed(Excel.Worksheet curWorksheet, string target,
             UnderwaterCrossing underwaterCrossing)
        {            
            curWorksheet.Columns.ClearFormats();
            curWorksheet.Rows.ClearFormats();

            Excel.Range entryPoint = curWorksheet.Cells.
                Find(target, Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole);

            object[,] data = curWorksheet.
                Range[curWorksheet.Cells[entryPoint.Row + 1, entryPoint.Column + 1],
                      curWorksheet.Cells[curWorksheet.UsedRange.Rows.Count, 
                      curWorksheet.UsedRange.Columns.Count]].Cells.Value;

            IEnumerable<DataRow> rezult = ArrayToDataTable(data).Select().
                Where(row => row.Field<string>(0) == "русло").
                Where(row => row.ItemArray.Count(cell => cell != DBNull.Value) > 1);

            if (!rezult.Any()) return;

            System.Data.DataTable dataTable = rezult.CopyToDataTable();

            Action<int> CheckValuesInColumn = delegate (int columnIndex)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    row[columnIndex] = GetFloatValue(row[columnIndex].ToString());
                }
            };

            string? value;
            switch (target)
            {
                case "недозаглубления":
                    CheckValuesInColumn(4);
                    DataRow[] row = dataTable.Select("[4] = MIN ([4])");
                    value = row[0].Field<string>(1);
                    underwaterCrossing.DeviationsRivebed!.LengthNGZ = value == null ? 0 : float.Parse(value);
                    break;
                case "оголения":
                    CheckValuesInColumn(4);
                    row = dataTable.Select("[4] = MAX ([4])");
                    value = row[0].Field<string>(1);
                    underwaterCrossing.DeviationsRivebed!.LengthDenudation = value == null ? 0 : float.Parse(value);
                    value = row[0].Field<string>(4);
                    underwaterCrossing.DeviationsRivebed!.MaxDepthDenudation = value == null ? 0 : float.Parse(value);
                    break;
                case "провисы":
                    CheckValuesInColumn(1);
                    row = dataTable.Select("[1] = MAX ([1])");
                    value = row[0].Field<string>(1);
                    underwaterCrossing.DeviationsRivebed!.LengthSag = value == null ? 0 : float.Parse(value);
                    underwaterCrossing.DeviationsRivebed.MaxLengthSinglePart = value == null ? 0 : float.Parse(value);
                    break;
                default:
                    break;
            }
        }
        /// <summary>
        /// Преобразование массива в объект DataTable
        /// </summary>
        /// <returns></returns>
        internal System.Data.DataTable ArrayToDataTable(object[,] array)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();  
            
            for (int i = 0; i < array.GetLength(1); i++)
                dataTable.Columns.Add(i.ToString(), typeof(string));

            for (int i = 0; i < array.GetLength(0); i++)
            {
                DataRow row = dataTable.NewRow();
                for (int j = 0; j < array.GetLength(1); j++)
                {
                    row[j.ToString()] = array[i + 1, j + 1];
                }
                dataTable.Rows.Add(row);
            }
            return dataTable;
        }
        /// <summary>
        /// Положение МТ по отношению к ППРР
        /// </summary>
        internal string GetPositionPipelineFor(DeviationsPVP deviationsPVP)
        {
            PropertyInfo[] props = typeof(DeviationsPVP).GetProperties();

            bool rezult = props.Where(prop => prop.Name.Contains("Riverbed")).
                               Any(prop => prop.GetValue(deviationsPVP)?.ToString() != "0");
            if (rezult)
                return "выше";
            else
                return "ниже";
        }
        /// <summary>
        /// Получение широты и долготы в десятичном формате
        /// </summary>
        internal string? ConvertGeoDegreeToDouble(string? coordinate, string pathFile, Excel.Range cell)
        {
            if (coordinate == null) return null;
            if (coordinate.Any(symb => char.IsLetter(symb)))
            {
                Loger.Log(pathFile, cell, $"Проверьте значение: {coordinate}");
            }

            coordinate = coordinate.Replace(".", ",");

            if (coordinate.IndexOf(",") == 2 || coordinate.IndexOf(",") == 3)
                return coordinate;
            else
            {
                coordinate = coordinate.Substring(0, coordinate.IndexOf(","));
                string[] geoValues = coordinate.Split(' ');

                double rezult =
                    double.Parse(geoValues[0]) +
                    double.Parse(geoValues[1]) / 60 +
                    double.Parse(geoValues[2]) / 3600;

                return Math.Round(rezult, 6).ToString();
            }
        }
        /// <summary>
        /// Получение номеров колонок шапки таблицы
        /// </summary>
        /// <param name="columns">Шаблон шапки таблицы</param>
        /// <exception cref="ArgumentException"></exception>
        internal void GetHeaderColumns(string pathFile, 
            Dictionary<string, List<int>> tableHeader, string sheetName)
        {
            Excel.Workbook curWorkbook;
            Excel.Worksheet curWorksheet;

            try
            {
                // Создание текущей рабочей книги
                curWorkbook = ExcelApp.Run.Workbooks.Open(pathFile);
                // Создание текущего ребочего листа
                curWorksheet = curWorkbook.Worksheets[sheetName];

                if (curWorksheet.AutoFilter != null && curWorksheet.AutoFilterMode == true)
                    curWorksheet.AutoFilter.ShowAllData();
            }
            catch (COMException)
            {
                throw new COMException($"Проверьте соответствие названия рабочего листа таблицы \n" +
                    $"\tфайл: {Path.GetFileName(pathFile)}   лист: {sheetName}");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            //Добавление номеров колонок в шаблон
            Action <Excel.Range, KeyValuePair<string, List<int>> > AddCollumnsFor = 
                delegate (Excel.Range cell, KeyValuePair<string, List<int>>column)
                {
                    if (cell == null)
                    {
                        Loger.Log(pathFile, $"Имя: {Path.GetFileName(pathFile)} - колонка: {column.Key} не найдена. ");
                        throw new ArgumentException($"колонка: {column.Key} не найдена. Проверьте таблицу");
                    }

                    for (int i = 0; i < cell.MergeArea.Columns.Count; i++)
                    {
                        column.Value.Add(cell.MergeArea.Column + i);
                    }
                };

            Excel.Range? cell = null;
            Excel.Range? otherSameCell = null;

            // Получение адресов колонок шапки таблицы
            foreach (var column in tableHeader)
            {
                cell = curWorksheet.Cells.Find(
                column.Key,
                Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlPart);

                if (cell == null)
                {
                    Loger.Log(pathFile, $"Имя: {Path.GetFileName(pathFile)} - колонка: {column.Key} не найдена. ");
                    throw new ArgumentException($"колонка: {column.Key} не найдена. Проверьте таблицу");
                }

                if (column.Key == "Наличие отклонений ПВП в русле")
                    otherSameCell = curWorksheet.Cells.FindNext(cell);

                if (otherSameCell != null)
                    AddCollumnsFor(otherSameCell, column);
                else
                    AddCollumnsFor(cell, column);   
            }

            // Закрытие текущей книги
            curWorkbook.Close(true);
        }
        public Table CheckTypeOfTable(string? sourcePathFile)
        {
            ArgumentNullException.ThrowIfNull(sourcePathFile);
            
            Excel.Workbook curWorkbook;
            Excel.Worksheet curWorksheet;

            try
            {
                // Создание текущей рабочей книги
                curWorkbook = ExcelApp.Run.Workbooks.Open(sourcePathFile);
                // Создание текущего ребочего листа
                curWorksheet = curWorkbook.Worksheets["Накопительная таблица"];
            }
            catch (COMException)
            {
                return Table.MeasuresTable;
            }
            catch (Exception ex)
            {
                Loger.Log(sourcePathFile, ex.Message);
                throw new Exception(ex.Message);
            }

            // Закрытие текущей книги
            curWorkbook.Close(true);
            
            return Table.AcumulativeTable;
        }
    }
}
