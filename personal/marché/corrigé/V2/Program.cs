using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;

Console.WriteLine("Place du marché");

ISheet sheet;
using (var stream = new FileStream("Place du marché.xlsx", FileMode.Open))
{
    stream.Position = 0;
    XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
    sheet = xssWorkbook.GetSheetAt(1);//select 2nd sheet

    var data = new List<MarketInfo>();

    for (int i = (sheet.FirstRowNum + 1/*skip header*/); i <= sheet.LastRowNum; i++)
    {
        IRow row = sheet.GetRow(i);
        int j = 0;
        data.Add(new MarketInfo(
            row.GetCell(j++).NumericCellValue.ToString(),
            row.GetCell(j++).StringCellValue,
            row.GetCell(j++).ToString(),
            row.GetCell(j++).ToString(),
            row.GetCell(j++).ToString(),
            row.GetCell(j++).ToString()
            ));
    }

    //Console.WriteLine(data);

    MarketInfo maxWatermelon= null;
    var peachSellers = new Dictionary<string,bool>();
    foreach(var item in data)
    {
        if(item.Product== "Pêches")
        {
            if(!peachSellers.ContainsKey(item.Seller))
            {
                peachSellers.Add(item.Seller, true);
            }
        }

        if(item.Product=="Pastèques")
        {
            if(maxWatermelon==null)
            {
                maxWatermelon = item;
            }
            else if(item.Quantity>maxWatermelon.Quantity)
            {
                maxWatermelon = item;
            }
        }
    }

    Console.WriteLine("NORMAL");
    Console.WriteLine($"Il y a {peachSellers.Keys.Count} vendeurs de pêches");
    Console.WriteLine($"C'est {maxWatermelon.Seller} qui a le plus de pastèques (stand {maxWatermelon.Place}, {maxWatermelon.Quantity} pièces)");







    //VERSION LINQ
    var peachProducersLINQ = (from item in data where item.Product == "Pêches" select 1).Count();

    var maxWatermelonSellerLINQ = (from item in data where item.Product == "Pastèques" where item.Quantity == data.Max(x => x.Quantity) select item).First();

    Console.WriteLine("LINQ");
    Console.WriteLine($"Il y a {peachProducersLINQ} vendeurs de pêches");
    Console.WriteLine($"C'est {maxWatermelonSellerLINQ.Seller} qui a le plus de pastèques (stand {maxWatermelonSellerLINQ.Place}, {maxWatermelonSellerLINQ.Quantity} pièces)");

}



class MarketInfo
{
    public MarketInfo(string place, 
        string seller, string product, string quantity, string unit, string pricPerUnit)
    {
        Place = place;
        Seller = seller;
        Product = product;
        Quantity = Convert.ToInt32(quantity);
        Unit = UnitConverter.fromString(unit);
        PricPerUnit = Convert.ToSingle(pricPerUnit);
    }

    public string Place { get; private set; }
    public string Seller { get; private set; }
    public string Product { get; private set; }
    public int Quantity { get; private set; }

    public Unit Unit { get; private set;}

    public float PricPerUnit { get; private set; }
}



enum Unit
{
    kg,unit,bag
}

static class UnitConverter
{
    public static Unit fromString(string unit)
    {
        switch(unit)
        {
            case "kg":return Unit.kg;
            case "pièce":return Unit.unit;
            case "sac": return Unit.bag;

            default: return Unit.kg;
        }
    }
}

