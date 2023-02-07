using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
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
        public string HeaderFontName = "Calibri";

        public int HeaderFontSize = 14;

        public bool setBold = true;

        public int rowContainingHeader = 2;

        public int firstColumnOfHeader = 1;

        [JsonIgnore]
        public XLColor HeaderBackgroundColor
        {
            get
            {
                if (HeaderBackgroundColor_ShadePercent == 0)
                    return XLColor.FromTheme(HeaderBackgroundColor_Theme);
                else
                    return XLColor.FromTheme(HeaderBackgroundColor_Theme, HeaderBackgroundColor_ShadePercent);
            }
        }
        [JsonConverter(typeof(StringEnumConverter))]
        public XLThemeColor HeaderBackgroundColor_Theme = XLThemeColor.Accent4;
        public double HeaderBackgroundColor_ShadePercent = 0.6;
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
        public string workbookTheme = "";

        public double maxColumnWidth = 50;

        public string defaultWorkbookName = "workbook_{0}.xlsx";

    }
}
