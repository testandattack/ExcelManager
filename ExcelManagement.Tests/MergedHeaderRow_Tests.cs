using System;
using Xunit;
using GTC.Extensions;
using System.Threading;
using System.Collections.Generic;
using Serilog;
using Xunit.Abstractions;
using System.Data;
using ClosedXML.Excel;
using ExcelManagement;
using System.IO;

namespace ExcelManagement.Tests
{
    // https://andrewlock.net/creating-strongly-typed-xunit-theory-test-data-with-theorydata/

    [Collection("xUnitHelper Collection")]
    public class MergedHeaderRow_Tests : IClassFixture<xUnitClassFixture>
    {
        xUnitCollectionFixture _collectionFixture;
        xUnitClassFixture _classFixture;
        private static string _logOutput = @"c:\temp\TestMergedHeaderRow";
        private ExcelBuilder _excelBuilder;
        private string comparisonSpreadsheet = "TestData\\MergedHeaderTestSpreadsheet.xlsx";

        public MergedHeaderRow_Tests(xUnitClassFixture classFixture, xUnitCollectionFixture collectionFixture, ITestOutputHelper output)
        {
            _collectionFixture = collectionFixture;
            _classFixture = classFixture;
            collectionFixture.ConfigureLogging(output, _logOutput);
            _excelBuilder = new ExcelBuilder();
        }

        [Theory]
        [InlineData("MergedHeaderRow", 1, 4, 2, 2, 14, true, XLThemeColor.Accent4, 0.6, true)]
        public void CreateMergedHeaderRow_Test(string title, int firstCellColumn, int lastCellColumn, int firstRowNumber, int lastRowNumber, int fontSize, bool setBold, XLThemeColor bgColor, double bgShadePercent, bool shouldPass)
        {
            try
            {
                // Arrange
                IXLWorksheet sheet1 = _excelBuilder.workbook.AddWorksheet("MergedHeaderRow");
                _excelBuilder.excelConfig.MergedCells.HeaderFontSize = fontSize;
                _excelBuilder.excelConfig.MergedCells.setBold = setBold;
                _excelBuilder.excelConfig.MergedCells.HeaderBackgroundColor_Theme = bgColor;
                _excelBuilder.excelConfig.MergedCells.HeaderBackgroundColor_ShadePercent = bgShadePercent;

                // Act
                _excelBuilder.CreateMergedHeaderRow(sheet1, title, firstCellColumn, lastCellColumn, firstRowNumber, lastRowNumber);
                _excelBuilder.Save();

                // Assert
                string sMessage;
                bool areTheSame = ClosedXML.Tests.ExcelDocsComparer.Compare(_excelBuilder.workbookName, comparisonSpreadsheet, out sMessage);
                Log.Information("ExcelComparer Output for {left} vs. {right}:\r\n{message}", _excelBuilder.workbookName, comparisonSpreadsheet, sMessage);
                Assert.Equal(shouldPass, areTheSame);
            }
            finally
            {
                if(File.Exists(_excelBuilder.workbookName))
                {
                    File.Delete(_excelBuilder.workbookName);
                }
            }
        }


    }
}
