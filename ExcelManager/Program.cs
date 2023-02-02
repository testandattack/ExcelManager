using System;
using ExcelManagement;
using System.IO;
using Newtonsoft.Json;

namespace ExcelManager
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelConfig excelConfig = new ExcelConfig();
            string configFile = JsonConvert.SerializeObject(excelConfig, Formatting.Indented);
            Console.WriteLine("Hello World!");
        }
    }
}
