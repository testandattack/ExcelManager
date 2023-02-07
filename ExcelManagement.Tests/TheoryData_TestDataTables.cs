using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelManagement.Tests
{
    public class TheoryData_TestDataTables : TheoryData<DataTable>
    {
        public TheoryData_TestDataTables()
        {
            DataTable myData = new DataTable();

            #region -- Column Mappings -----
            // 1 Id
            // 2 ItemId
            // 3 Museum Piece?
            // 4 Box
            // 5 Qty
            // 6 Name
            // 12 Retail
            // 14 Last Updated
            #endregion

            #region -- Columns -----
            myData.Columns.Add(new DataColumn("Id", typeof(System.Int32)));
            myData.Columns.Add(new DataColumn("ItemId", typeof(System.String)));
            myData.Columns.Add(new DataColumn("Museum Piece?", typeof(System.Boolean)));
            myData.Columns.Add(new DataColumn("Box", typeof(System.String)));
            myData.Columns.Add(new DataColumn("Qty", typeof(System.Int32)));
            myData.Columns.Add(new DataColumn("Name", typeof(System.String)));
            myData.Columns.Add(new DataColumn("Retail", typeof(System.Double)));
            myData.Columns.Add(new DataColumn("Last Updated", typeof(System.DateTime)));
            #endregion

            #region -- Data -----
            myData.Rows.Add(1, "7194-1", true, "C01", 1, "Yoda", 23.4, "12/20/2022");
            myData.Rows.Add(2, "8458-1", true, "C01", 1, "Silver Champion", 324.45, "12/20/2022");
            myData.Rows.Add(3, "42009-1", true, "C01", 1, "Mobile Crane MkII", 1.2, "12/20/2022");
            myData.Rows.Add(4, "42082-1", true, "C01", 1, "Rough Terrain Crane", 94.1500000000001, "12/20/2022");
            myData.Rows.Add(5, "76042-1", true, "C02", 1, "SHIELD Helicarrier", 83.0500000000001, "12/20/2022");
            myData.Rows.Add(6, "001-1", true, "Ch01", 1, "Gears", 71.9500000000001, "12/20/2022");
            myData.Rows.Add(7, "112-2", true, "Ch01", 1, "Train with Motor", 60.8500000000002, "12/20/2022");
            myData.Rows.Add(8, "8437-1", true, "Ch01", 1, "Future Car", 49.7500000000002, "12/20/2022");
            myData.Rows.Add(9, "8700-1", true, "Ch01", 1, "Expert Builder Power Pack", 38.6500000000002, "12/20/2022");
            myData.Rows.Add(10, "8858-1", true, "Ch01", 1, "Auto Engines", 27.5500000000003, "12/20/2022");
            myData.Rows.Add(11, "8865-1", true, "Ch01", 1, "Test Car", 16.45, "12/20/2022");
            myData.Rows.Add(12, "42043-1", true, "Ch01", 1, "Mercedes Arocs", 5.34999999999997, "12/20/2022");
            myData.Rows.Add(13, "kabrobo-1", true, "Ch01", 1, "Robo Riders 4 pack", -5.75000000000003, "12/20/2022");
            myData.Rows.Add(14, "5541-1", true, "Ch02", 1, "Blue Fury", -16.8500000000001, "12/20/2022");
            myData.Rows.Add(15, "5542-1", true, "Ch02", 1, "Black Thunder", -27.95, "12/20/2022");
            myData.Rows.Add(16, "5561-1", true, "Ch02", 1, "Big Foot 4x4", -39.05, "12/20/2022");
            myData.Rows.Add(17, "5590-1", true, "Ch02", 1, "Whirl and Wheel Super Truck", -50.15, "12/20/2022");
            myData.Rows.Add(18, "8459-1", true, "Ch02", 1, "Pneumatic Front End Loader", -61.25, "12/20/2022");
            #endregion

            Add(myData);
        }
    }
}
