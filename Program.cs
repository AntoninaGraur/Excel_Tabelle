


using System;
using System.Data;
using System.IO;
using ExcelDataReader;


class Programm
{
    static void Main(string[] args)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        string excelFilePath = "C:/UsersA.Graur/Documents/BKU/SWD/Bestellungen.xlsx";
        string csvFilePath = "C:/UsersA.Graur/Documents/BKU/SWD/BestellungenCSV.xlsx"; 

        using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
        {
            using ( var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet();

                DataTable table = result.Table
            }
        }
    }
}