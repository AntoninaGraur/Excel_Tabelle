


using System;
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
                Console.WriteLine("Inhalt der Excel-Datei: ");
                foreach (DataRow  row in table.Rows)
                {
                    for (int i=0; i< table.Columns.Count; i++)
                    {
                        Console.Write(row[i].ToString() + ";  ");
                    }
                    Console.WriteLine("                     ");
                } Console.WriteLine("In CSV umgewandelt.");

                  FindenTeuerstenBilligstenArtikel(table);
            }
           
        }
    }

  static void FindenTeuerstenBilligstenArtikel(DataTable table)
    {
        string teuerstenArtikel = "";
        string billigstenArtikel = "";

        double maxPrice = double.MinValue;
        double minPrice = double.MaxValue;

        foreach (DataRow row in table.Rows)
        {
            string artikel = row["Artikel"].ToString();
            string priceText = row["Nettopreis"].ToString().Replace("€", "").Replace(".", "").Replace(",", ".").Trim();

            if (double.TryParse(priceText, out double price))
            {
                if (price>maxPrice)
                {
                    maxPrice = price;
                    teuerstenArtikel = artikel;
                }

                if (price < minPrice)
                {
                    minPrice = price;
                    billigstenArtikel = artikel;
                }
            }
        }
        Console.WriteLine($"Teuerster Artikel : {teuerstenArtikel}, Kostet: {maxPrice}€ ");
        Console.WriteLine("///////////////////////////////////////////////////////////////");
        Console.WriteLine($"Teuerster Artikel : {billigstenArtikel}, Kostet: {minPrice}€ ");

    }
}