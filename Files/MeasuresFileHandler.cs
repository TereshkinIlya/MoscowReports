using System.Runtime.InteropServices;
using Files.Abstracts;
using Files.Tables;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Files
{
    public enum Table
    {
        AcumulativeTable, MeasuresTable
    };
    public class MeasuresFileHandler : FileHandler
    {
        private object SourceTable {  get; set; }
        private MeasuresTable ReceiverTable { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public MeasuresFileHandler(object? sourceTable, MeasuresTable measuresTable, 
            IProgress<object[]> progress)
        {
            ArgumentNullException.ThrowIfNull(sourceTable);
            ArgumentNullException.ThrowIfNull(measuresTable);
            ArgumentNullException.ThrowIfNull(progress);

            ReceiverTable = measuresTable;
            Progress = progress;
            SourceTable = sourceTable;

        } 
        public override void HandleFiles()
        {
            FillTable();
        }
        private void FillTable()
        {
            try
            {
                if (SourceTable.GetType() == typeof(AcumulativeTable))
                    HandleAsAcumalative();
                else
                    HandleAsMeasures();

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
        /// Перенос данных ( 13 параметров, ОТС) в мероприятия из накопительной
        /// </summary>
        /// <exception cref="ArgumentNullException"></exception>
        private void HandleAsAcumalative()
        {
            int indexProgress = 1;
            int quantRows = 1;
            
            var acumulativeTable = SourceTable as AcumulativeTable ??
                    throw new ArgumentNullException(nameof(SourceTable));

            Excel.Workbook sourceTable = ExcelApp.Run.Workbooks.Open(acumulativeTable.Path);
            Excel.Workbook receiverTable = ExcelApp.Run.Workbooks.Open(ReceiverTable.Path);

            Excel.Worksheet sourceWorksheet = sourceTable.Worksheets["Накопительная таблица"];
            Excel.Worksheet receiverWorksheet = receiverTable.Worksheets["Сводная"];

            Excel.Range sourceID;
            Excel.Range receiverID;

            Excel.Range measuresIdCell = receiverWorksheet.Cells.
                Find("ID", Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole);

            Excel.Range acumulativeIdCell = sourceWorksheet.Cells.
                Find("ID", Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole);

            Excel.Range measuresIdColumn = receiverWorksheet.Range
                [
                receiverWorksheet.Cells[measuresIdCell.Row + 3, measuresIdCell.Column],
                receiverWorksheet.Cells[receiverWorksheet.UsedRange.Rows.Count, measuresIdCell.Column]
                ];

            if (sourceWorksheet.AutoFilter != null && sourceWorksheet.AutoFilterMode == true)
                sourceWorksheet.AutoFilter.ShowAllData();

            if (receiverWorksheet.AutoFilter != null && receiverWorksheet.AutoFilterMode == true)
                receiverWorksheet.AutoFilter.ShowAllData();

            Excel.Range paramsRange = receiverWorksheet.Range
                [
                receiverWorksheet.Cells[measuresIdCell.Row + 3, ReceiverTable.TableColumns["Номер заключения"][0]],
                receiverWorksheet.Cells[receiverWorksheet.UsedRange.Rows.Count, ReceiverTable.TableColumns["ЭХЗ"][0]]
                ];

            paramsRange.Clear();
            quantRows = sourceWorksheet.UsedRange.Rows.Count;
            foreach (Excel.Range item in sourceWorksheet.UsedRange.Rows)
            {
                sourceID = sourceWorksheet.Cells[item.Row, acumulativeIdCell.Column];

                Progress.Report(new object[] { quantRows, indexProgress++, "Копирование параметров БЭ из накопительной...", 1 });

                try
                {
                    receiverID = measuresIdColumn.Cells.
                    Find(sourceID.Value, Type.Missing,
                    Excel.XlFindLookIn.xlValues,
                    Excel.XlLookAt.xlWhole);

                    if (receiverID == null) continue;

                    sourceWorksheet.Cells[sourceID.Row, acumulativeTable.TableColumns["Дата выдачи заключения"][0]].
                        Copy();
                    receiverWorksheet.Cells[receiverID.Row, ReceiverTable.TableColumns["Дата выдачи заключения"][0]].
                        PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats);

                    sourceWorksheet.Cells[sourceID.Row, acumulativeTable.TableColumns["Номер ОТС"][0]].
                        Copy();
                    receiverWorksheet.Cells[receiverID.Row, ReceiverTable.TableColumns["Номер заключения"][0]].
                        PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats);

                    sourceWorksheet.Range
                        [
                        sourceWorksheet.Cells[sourceID.Row, acumulativeTable.TableColumns["Проведение ВТД"][0]],
                        sourceWorksheet.Cells[sourceID.Row, acumulativeTable.TableColumns["Обследование корр"][0]]
                        ].
                        Copy();

                    receiverWorksheet.Range
                        [
                        receiverWorksheet.Cells[receiverID.Row, ReceiverTable.TableColumns["ВТД"][0]],
                        receiverWorksheet.Cells[receiverID.Row, ReceiverTable.TableColumns["ЭХЗ"][0]]
                        ].
                        PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats);
                }
                catch(COMException)
                {}
                catch (Exception)
                {}
            }
            paramsRange.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

            ExcelApp.Run.Workbooks[1].Save();
            ExcelApp.Run.Workbooks[2].Save();
            ExcelApp.Run.Workbooks.Close();
        }
        /// <summary>
        /// Перенос меропритий ОСТов в итоговую таблицу мероприятий
        /// </summary>
        /// <exception cref="ArgumentNullException"></exception>
        private void HandleAsMeasures()
        {
            int indexProgress = 1;
            int quantRows = 0;

            var measuresTable = SourceTable as MeasuresTable ??
                    throw new ArgumentNullException(nameof(SourceTable));

            Excel.Workbook sourceTable = ExcelApp.Run.Workbooks.Open(measuresTable.Path);
            Excel.Workbook receiverTable = ExcelApp.Run.Workbooks.Open(ReceiverTable.Path);

            Excel.Worksheet sourceWorksheet = sourceTable.Worksheets["Сводная"];
            Excel.Worksheet receiverWorksheet = receiverTable.Worksheets["Сводная"];

            Excel.Range sourceID;
            Excel.Range receiverID;

            Excel.Range sourceIdCell = sourceWorksheet.Cells.
                Find("ID", Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole);

            Excel.Range receiverIdCell = receiverWorksheet.Cells.
                Find("ID", Type.Missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlWhole);

            Excel.Range measuresIdColumn = receiverWorksheet.Range
                [
                receiverWorksheet.Cells[receiverIdCell.Row, receiverIdCell.Column],
                receiverWorksheet.Cells[receiverWorksheet.UsedRange.Rows.Count, receiverIdCell.Column]
                ];

            if (receiverWorksheet.AutoFilter != null && receiverWorksheet.AutoFilterMode == true)
                receiverWorksheet.AutoFilter.ShowAllData();

            if (sourceWorksheet.AutoFilter != null &&
                sourceWorksheet.AutoFilter.Filters[measuresTable.TableColumns["ОСТ"][0] - 1].On == false)
                throw new ArgumentException("Сняты фильтры в мероприятиях ОСТа! Проверьте таблицу");

            Excel.Range filteredRows = sourceWorksheet.Range
                [
                sourceWorksheet.Cells[sourceIdCell.Row, measuresTable.TableColumns["№№ п/п"][0]],
                sourceWorksheet.Cells[sourceWorksheet.UsedRange.Rows.Count, sourceIdCell.Column]
                ].SpecialCells(XlCellType.xlCellTypeVisible, Type.Missing);


            for (int areaIndex = 1; areaIndex <= filteredRows.Areas.Count; areaIndex++)
            {
                quantRows += filteredRows.Areas[areaIndex].Rows.Count;
            }


            for (int areaIndex = 1; areaIndex <= filteredRows.Areas.Count; areaIndex++)
            {
                for (int row = 1; row <= filteredRows.Areas[areaIndex].Rows.Count; row++)
                {
                    sourceID = sourceWorksheet.Cells[filteredRows.Areas[areaIndex].Rows[row].Row, sourceIdCell.Column];
                    
                    Progress.Report(new object[] { quantRows, indexProgress++, "Копирование мероприятий из таблицы ОСТа...", 1 });

                    try
                    {
                        receiverID = measuresIdColumn.Cells.
                        Find(sourceID.Value, Type.Missing,
                        Excel.XlFindLookIn.xlValues,
                        Excel.XlLookAt.xlWhole);

                        if (receiverID == null) continue;

                        var tes = sourceID.Value;
                        sourceWorksheet.Range
                            [
                            sourceWorksheet.Cells[sourceID.Row, ReceiverTable.TableColumns["WM"][0]],
                            sourceWorksheet.Cells[sourceID.Row, ReceiverTable.TableColumns["CDL"][0]]
                            ].
                            Copy();

                        receiverWorksheet.Range
                            [
                            receiverWorksheet.Cells[receiverID.Row, ReceiverTable.TableColumns["WM"][0]],
                            receiverWorksheet.Cells[receiverID.Row, ReceiverTable.TableColumns["CDL"][0]]
                            ].
                            PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats);

                        sourceWorksheet.Cells[sourceID.Row, ReceiverTable.TableColumns["Мероприятия"][0]].
                            Copy();
                        receiverWorksheet.Cells[receiverID.Row, ReceiverTable.TableColumns["Мероприятия"][0]].
                            PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats);

                    }
                    catch (COMException)
                    { }
                    catch (Exception)
                    { }
                }
            }

            receiverWorksheet.UsedRange.AutoFilter(ReceiverTable.TableColumns["Дата выдачи заключения"][0] - 2, "<>");

            ExcelApp.Run.Workbooks[1].Save();
            ExcelApp.Run.Workbooks[2].Save();
            ExcelApp.Run.Workbooks.Close();
        }
    }
}
