using System;
using System.Data;
using System.Linq;
using ClosedXML.Excel;


                                   //Ultra//

class Program
{
    static void Main()
    {
        string path = @"C:\Users\pn25kdv\Documents\GitHub\323-Programmation_fonctionnelle\personal\marché\Place_du_marché.xlsx";

        string keywordPeches = "Pêches";
        string keywordPasteques = "Pastèques";

        DataTable dt = new DataTable();

        using (var workbook = new XLWorkbook(path))
        {
            var worksheet = workbook.Worksheet(2); // 2-й лист

            bool firstRow = true;
            foreach (var row in worksheet.RowsUsed())
            {
                if (firstRow) // первая строка = заголовки
                {
                    foreach (var cell in row.Cells())
                        dt.Columns.Add(cell.Value.ToString());
                    firstRow = false;
                }
                else
                {
                    dt.Rows.Add();
                    int i = 0;
                    foreach (var cell in row.Cells())
                        dt.Rows[dt.Rows.Count - 1][i++] = cell.Value.ToString();
                }
            }
        }

        // Сколько продавцов торгуют персиками
        int nmbr_peches_seller = dt.AsEnumerable()
            .Count(r => r["Produit"].ToString().Contains(keywordPeches, StringComparison.OrdinalIgnoreCase));

        Console.WriteLine($"Il y a {nmbr_peches_seller} vendeurs de pêches");

        // Кто продаёт больше всего арбузов
        var pastequesGroup = dt.AsEnumerable()
            .Where(r => r["Produit"].ToString().Contains(keywordPasteques, StringComparison.OrdinalIgnoreCase))
            .GroupBy(r => r["Producteur"].ToString()) // сгруппировать по продавцу
            .Select(g => new
            {
                Vendeur = g.Key,
                Stand = g.First()["Emplacement"].ToString(),
                Total = g.Sum(r => int.Parse(r["Quantité"].ToString())),
               
            })
            .OrderByDescending(x => x.Total)
            .FirstOrDefault();

        if (pastequesGroup != null)
        {
            Console.WriteLine($"C'est {pastequesGroup.Vendeur} qui a le plus de pastèques (stand {pastequesGroup.Stand}, {pastequesGroup.Total} pièces)");
        }
    }
}



                                    //Mine//
/*
class Program
{
   static void Main()
    {
        string path = @"C:\Users\pn25kdv\Documents\GitHub\323-Programmation_fonctionnelle\personal\marché\Place du marché.xlsx";

        int seller_most_pasteques;
        int seller_place;

        string keywordPeches = "Pêches";
        string keywordPasteques = "Pastèques";

        DataTable dt = new DataTable();

        using (var workbook = new XLWorkbook(path))
        {
            var worksheet = workbook.Worksheet(2);


            int nmbr_peches_seller = dt.AsEnumerable()
           .Count(r => r["Produit"].ToString().Contains(keywordPeches, StringComparison.OrdinalIgnoreCase));

            int nmbr_pasteques_seller = dt.AsEnumerable()
           .Count(r => r["Produit"].ToString().Contains(keywordPasteques, StringComparison.OrdinalIgnoreCase));


            Console.WriteLine($"Il y a {nmbr_peches_seller} vendeurs de pêches");
            Console.WriteLine($"C'est {seller_most_pasteques} qui a le plus de pastèques (stand {seller_place}, {nmbr_pasteques_seller} pièces)");
        }
    }
}*/