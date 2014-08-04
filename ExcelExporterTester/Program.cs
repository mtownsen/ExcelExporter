using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using CreateExcelSpreadsheet;
using System.Web;

namespace ExcelExporterTester
{
    class Program
    {
        static void Main(string[] args)
        {
            //Instantiate a instance of our ExcelHelper class to allow us to access the excel export
            ExcelHelper excelHelper = new ExcelHelper();

            //Now lets generate some randome test data to push through to the excel Creator
            ExcelExportTestData excelExportDataGenerator = new ExcelExportTestData();
            List<ExcelExporterTester.ExcelExportTestData.TestData> exportData = excelExportDataGenerator.GenerateTestData();

            //Now that we have valid data lets pass it into the excel creator with:
                //The path to store the file to.
                //The test data.
                //The name for the spreadsheet tab.
                //And optionally the list of column header names
            excelHelper.Create<ExcelExporterTester.ExcelExportTestData.TestData>(
                "c:/users/thunderfan/documents/visual studio 2012/Projects/ExcelExporter/ExcelExporterTester/TestExcelFiles/TestExcelExport.xlsx", exportData, "TestExportData", null);

        }
    }
}
