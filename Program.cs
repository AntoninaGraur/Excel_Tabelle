


using System;
using System.Collections.Immutable;
using System.Data;
using System.IO;
using ExcelDataReader;


class Programm
{
    static void Main(string[] args)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        string excelFilePath = "C:/Users/A.Graur/Documents/BKU/SWD/Bestellungen.xlsx";
        string csvFilePath = "C:/Users/A.Graur/Documents/BKU/SWD/BestellungenCSV.csv"; 

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
                        if( i < table.Columns.Count -1) writer.Write(",");
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
                //Datei auf dem Bildschirm ausgeben
                Console.WriteLine("    Inhalt der Excel-Datei  ");
                foreach (DataRow  row in table.Rows)
                {
                    for (int i=0; i< table.Columns.Count; i++)
                    {
                        Console.Write(row[i].ToString() +  ";  ");
                    }
                    Console.WriteLine("                     ");
                } Console.WriteLine("****************In CSV umgewandelt******************");

                
            }
           
        }
        GetArtikel();
    }

    static void GetArtikel()
    {
        Console.WriteLine("//******************Artikel + Nettopreis*********//");


        string filePath = @"C:\Users\A.Graur\Documents\BKU\SWD\Bestellungen.xlsx";

        List<string> list = new List<string>();

        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
            {
                int counter = 0;

                while (reader.Read())
                {
                    counter++;

                    if (counter > 1)
                    {
                        string artikel = reader.GetValue(8)?.ToString() ?? "Falls";
                        string nettopreis = reader.GetValue(9)?.ToString() ?? "";

                        string fullSTR = $"Artikel: {artikel} -- Nettotpreis: {nettopreis}";
                        list.Add(fullSTR);
                    }

                }
            }
        }

        foreach (var item in list)
        {
            Console.WriteLine(item);
        }

        List<string> nameList = new List<string>();


        Console.ReadKey();
    }
}

