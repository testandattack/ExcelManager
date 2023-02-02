using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using ClosedXML;
using ClosedXML.Excel;
using GTC.Extensions;
using Newtonsoft.Json;

namespace ExcelManagement
{
    public class ExcelBuilder
    {
        private ExcelConfig _excelConfig;

        public XLWorkbook workbook { get; set; }
        
        public string workbookName { get; set; }

        #region -- Constructors -----
        public ExcelBuilder()
        {
            workbook = new XLWorkbook();
        }

        public ExcelBuilder(string filename)
        {
            workbook = new XLWorkbook();
            Initialize(filename);
        }

        private void Initialize(string filename)
        {
            using (StreamReader sr = new StreamReader("ExcelConfig.json"))
            {
                _excelConfig = JsonConvert.DeserializeObject<ExcelConfig>(sr.ReadToEnd());
            }

            try
            {
                if (filename == "")
                {
                    string defaultFileName = _excelConfig.General.defaultWorkbookName;
                    if (defaultFileName.Contains("{0}"))
                    {
                        filename = String.Format(defaultFileName, DateTime.Now.ToShortDateString());
                    }
                    else
                    {
                        filename = defaultFileName;
                    }
                }

                if (filename.EndsWith(".xlsx") == true)
                    workbookName = filename;
                else
                    workbookName = $"{filename}.xlsx";
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }
        #endregion

        #region -- AddWorksheetFromDataTable -----
        public IXLWorksheet AddWorksheetFromDataTable(string name, DataTable dataTable, string tableName)
        {
            return AddWorksheetFromDataTable(name, dataTable, tableName, "", "");
        }

        public IXLWorksheet AddWorksheetFromDataTable(string name, DataTable dataTable, string tableName, string columnToTextWrap)
        {
            var ws = workbook.AddWorksheet(name);
            CreateMergedHeaderRow(ws, name, dataTable.Columns.Count);

            CreateExcelTableFromDataTable(dataTable, tableName, ws);

            if (_excelConfig.DataTables.useTextRotationOnHeaders)
            {
                var range1 = ws.Range(
                    ws.Cell(4, 1).Address,
                    ws.Cell(4, dataTable.Columns.Count).Address);
                range1.Style.Alignment.SetTextRotation(90);
                range1.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            }

            ws.SheetView.FreezeRows(4);
            ws.Columns().AdjustToContents();

            if (columnToTextWrap != "")
            {
                if (dataTable.Columns.Contains(columnToTextWrap))
                {
                    int columnIndex = dataTable.Columns[columnToTextWrap].Ordinal + 1;
                    ws.Column(columnIndex).Style.Alignment.SetWrapText(true);
                }
            }
            return ws;
        }
        #endregion

        #region -- CreateExcelTableFromDataTable -----
        public IXLTable CreateExcelTableFromDataTable(DataTable dataTable, string tableName, string sheetName)
        {
            var ws = workbook.AddWorksheet(sheetName);
            return CreateExcelTableFromDataTable(dataTable, tableName, ws, _excelConfig.DataTables.startingRow, _excelConfig.DataTables.startingColumn);
        }

        public IXLTable CreateExcelTableFromDataTable(DataTable dataTable, string tableName, IXLWorksheet ws)
        {
            return CreateExcelTableFromDataTable(dataTable, tableName, ws, _excelConfig.DataTables.startingRow, _excelConfig.DataTables.startingColumn);
        }

        public IXLTable CreateExcelTableFromDataTable(DataTable dataTable, string tableName, IXLWorksheet ws, int startRow, int startColumn)
        {
            var newTable = ws.Cell(startRow, startColumn).InsertTable(dataTable, tableName, true);
            newTable.Theme = _excelConfig.DataTables.tableTheme;
            return newTable;
        }
        #endregion

        #region -- CreateMergedHeaderRow -----
        public void CreateMergedHeaderRow(IXLWorksheet sheet, string title, int numColumns)
        {
            var range = sheet.Range(sheet.Cell(1, 1).Address, sheet.Cell(1, numColumns).Address);
            range.Merge();
            range.Style.Font.SetBold().Font.FontSize = _excelConfig.MergedCells.HeaderFontSize;
            range.Style.Fill.SetBackgroundColor(_excelConfig.MergedCells.HeaderBackgroundColor);
            range.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Value = title;
        }

        public void CreateMergedHeaderRow(IXLWorksheet sheet, string title, int iFirstCellColumn, int iLastCellColumn)
        {
            var range = sheet.Range(2, iFirstCellColumn, 2, iLastCellColumn);
            range.Merge().Style.Font.SetBold().Font.FontSize = _excelConfig.MergedCells.HeaderFontSize;
            range.Style.Fill.SetBackgroundColor(_excelConfig.MergedCells.HeaderBackgroundColor);
            range.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Value = title;
        }
        #endregion

        #region -- SetConditionalColumnWidth -----
        public void SetConditionalColumnWidth(IXLWorksheet sheet, int column)
        {
            SetConditionalColumnWidth(sheet, column, _excelConfig.General.maxColumnWidth);
        }

        public void SetConditionalColumnWidth(IXLWorksheet sheet, int column, double maxColumnwidth)
        {
            if (sheet.Column(column).Width > maxColumnwidth)
            {
                sheet.Column(column).Width = maxColumnwidth;
                sheet.Column(column).Style.Alignment.WrapText = true;
            }
        }
        #endregion

        #region -- Misc -----
        public void AddThreeColumnDataTableComparisonSheet(IXLWorksheet sheet, DataTable data1, string table1Name, DataTable data2, string table2Name, int sortColumn = 0)
        {
            string table1column1 = data1.Columns[sortColumn].ColumnName;
            string table2column1 = data2.Columns[sortColumn].ColumnName;

            var table1 = sheet.Cell(3, 1).InsertTable(data1, table1Name);  // Table is 9 columns wide            
            CreateMergedHeaderRow(sheet, table1Name, 1, 3);
            table1.Sort(table1column1);

            var table2 = sheet.Cell(3, 4).InsertTable(data2, table2Name);  // Table is 9 columns wide            
            CreateMergedHeaderRow(sheet, table2Name, 4, 6);
            table2.Sort(table2column1);

            // start at the end row of the smaller data table
            int currentDataRow = data1.Rows.Count > data2.Rows.Count ? data2.Rows.Count : data1.Rows.Count;

            // loop through the two tables and add blank entries for each missing item.
            for (int x = currentDataRow - 1; x > -1; x--)
            {
                string value1 = table1.DataRange.Row(x).Field(table1column1).Value.ToString(); string value2 = table2.DataRange.Row(x).Field(table2column1).Value.ToString(); int comparison = value1.CompareTo(value2); if (comparison == -1)
                {
                    //insert row into table 1                    
                    table1.DataRange.Row(x).InsertRowsAbove(1);
                }
                else if (comparison == 1)
                {
                    //insert row into table 2                    
                    table2.DataRange.Row(x).InsertRowsAbove(1);
                }
                else
                {
                    //do nothing                    
                    continue;
                }
            }
            sheet.Columns().AdjustToContents();
            sheet.SheetView.FreezeRows(3);
        }

        public void Save()
        {
            try
            {
                workbook.SaveAs(workbookName);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving file. {ex.Message}");

            }
        }
        #endregion
    }
}
