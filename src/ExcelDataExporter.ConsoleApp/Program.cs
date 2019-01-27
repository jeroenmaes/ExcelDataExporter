using ExcelToXmlParser.Helpers;
using ExcelToXmlParser.Services;
using System;
using Common.Logging;
using System.IO;
using System.Net.Http;

namespace ExcelToXmlParser.ConsoleApp
{
    class Program
    {       
        static void Main(string[] args)
        {
            IExcelConverter converter = new ExcelConverter();
            IFileStorage fileStore = new FolderFileStorage();

            // Input Excel file
            var inputFilePath = @"C:\_git\ExcelDataExporter\tst\Resources\Financial Sample.xlsx";
            
            var inputStream = fileStore.GetFileStream(inputFilePath);

            // Load Stream
            converter.LoadStream(inputStream);

            // Convert Excel To XML                
            //var xmlString = converter.GetXML(");
            var xmlString = converter.GetXML("Excel", "RowItem", "http://sample.com");

            // Save Instance
            fileStore.SaveFile("Excel.xml", StreamHelpers.ConvertStringToStream(xmlString));

            // Convert Excel To JSON                
            var jsonString = converter.GetJSON();

            // Save Instance
            fileStore.SaveFile("Excel_JSON.json", StreamHelpers.ConvertStringToStream(jsonString));

            Console.ReadKey();
        }
    }
}
