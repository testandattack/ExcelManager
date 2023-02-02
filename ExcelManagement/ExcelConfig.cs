using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelManagement
{
    public class ExcelConfig
    {
        public MergedCells MergedCells { get; set; }

        public DataTables DataTables { get; set; }

        public General General { get; set; }

        public ExcelConfig()
        {
            MergedCells = new MergedCells();
            DataTables = new DataTables();
            General = new General();
        }
    }

    public class MergedCells
    {
        public int HeaderFontSize = 14;

        [JsonIgnore]
        public XLColor HeaderBackgroundColor;
        public System.Drawing.Color HeaderBackgroundColorAsSystemColor = XLColor.LightGray.Color;

        public MergedCells()
        {
            HeaderBackgroundColor = XLColor.FromColor(HeaderBackgroundColorAsSystemColor);
        }
    }

    public class DataTables
    {
        public int startingRow = 1;

        public int startingColumn = 1;

        [JsonIgnore]
        public XLTableTheme tableTheme;
        public string tableThemeName = XLTableTheme.TableStyleLight9.Name;

        public int textRotation = 90;

        public bool useTextRotationOnHeaders = true;

        public DataTables()
        {
            tableTheme = XLTableTheme.FromName(tableThemeName);
        }
    }

    public class General
    {
        public double maxColumnWidth = 50;

        public string defaultWorkbookName = "workbook_{0}.xlsx";

    }
}
