


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

                DataTable table = result.Tables[0];
                using (var writer = new StreamWriter(csvFilePath))
                {
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        writer.Write(table.Columns[i].ColumnName);
                        if( i < table.Columns.Count -1 )
                        {
                            writer.Write(",");
                        }
                        writer.WriteLine();

                    }

                    foreach (DataRow row in table.Rows)
                    {
                        for (int  i = 0;  i < table.Columns.Count;  i++)
                        {
                            writer.Write(row[i].ToString());
                            if (i < table.Columns.Count - 1) writer.WriteLine(",");
                        }
                        writer.WriteLine();
                    }
                }

            }
        }
        Console.WriteLine("In CSV umgewandelt.");
    }
}