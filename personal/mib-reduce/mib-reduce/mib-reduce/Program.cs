using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static mib_reduce.Program;

namespace mib_reduce
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Product> products = new List<Product>{
    new Product { Location = 1, Producer = "Bornand", ProductName = "Pommes", Quantity = 20,Unit = "kg", PricePerUnit = 5.50 },
    new Product { Location = 1, Producer = "Bornand", ProductName = "Poires", Quantity = 16,Unit = "kg", PricePerUnit = 5.50 },
    new Product { Location = 1, Producer = "Bornand", ProductName = "Pastèques", Quantity = 14,Unit = "pièce", PricePerUnit = 5.50 },
    new Product { Location = 1, Producer = "Bornand", ProductName = "Melons", Quantity = 5,Unit = "kg", PricePerUnit = 5.50 },
    new Product { Location = 2, Producer = "Dumont", ProductName = "Noix", Quantity = 20,Unit = "sac", PricePerUnit = 5.50 },
    new Product { Location = 2, Producer = "Dumont", ProductName = "Raisin", Quantity = 6,Unit = "kg", PricePerUnit = 5.50 },
    new Product { Location = 2, Producer = "Dumont", ProductName = "Pruneaux", Quantity = 13,Unit = "kg", PricePerUnit = 5.50 },
    new Product { Location = 2, Producer = "Dumont", ProductName = "roseilles", Quantity = 12,Unit = "kg", PricePerUnit = 5.50 },
    new Product { Location = 2, Producer = "Dumont", ProductName = "Groseilles", Quantity = 101,Unit = "kg", PricePerUnit = 5.50 },
            };

            // 0. Quantité de groseilles disponibles
            var groseillesQquantity = products.Where(p => p.ProductName.ToLower() == "groseilles").Sum(p => p.Quantity);
            Console.WriteLine($"Quantité de groseilles disponibles: {groseillesQquantity}");

            // 1. Chiffre d'affaire total par marchand
// Groupe les produits par producteur, calcule le chiffre d'affaires pour chaque groupe et matérialise le résultat en liste.
var ChiffreAffaire = products
    .GroupBy(p => p.Producer) // regroupe les éléments de 'products' par la propriété 'Producer'
    .Select(g => new          // pour chaque groupe, crée un objet anonyme contenant :
    {
        Producer = g.Key,     //   - la clé du groupe (la valeur par laquelle il est groupé- le nom du producteur)
        ChiffreAffaire = g.Sum(p => p.Quantity * p.PricePerUnit) //   - la somme des (Quantité * PrixUnitaire)
    })
    .ToList();                // exécute la requête et retourne List<...>

Console.WriteLine("\nChiffre d'affaire par marchand: ");
foreach (var item in ChiffreAffaire)
{
    // Affiche le producteur et son chiffre d'affaires
    Console.WriteLine($"{item.Producer}: {item.ChiffreAffaire} CHF");
}

            //2. Le plus grand, le plus petit et la moyenne de ces chiffres d’affaire
            var LPGCA = ChiffreAffaire.Max(x => x.ChiffreAffaire);
            var LPPCA = ChiffreAffaire.Min(x => x.ChiffreAffaire);
            var MYNCA = ChiffreAffaire.Average(x => x.ChiffreAffaire);
            Console.WriteLine($"\nLe plus grand:{LPGCA}CHF,le plus petit: {LPPCA}CHF,la moyenne: {MYNCA}CHF");

            //3.  Le marchand ayant le plus de noix à vendre
            
            Console.WriteLine($"\n  Le marchand ayant le plus de noix à vendre: {Producer}");

            //4. Le marchand ayant le plus d’affinités avec ses produits
            Console.WriteLine($"Le marchand ayant le plus d’affinités avec ses produits: {LPAFF}");


            int Affinity(string name, string product)
            {
                return name.GroupBy(letter => letter)
                    .Union(product.GroupBy(letter => letter))
                    .Sum(group => group.Count());
            }
        }
        public class Product {
            public int Location { get ;set;}
            public string Producer { get; set; }
            public string ProductName { get; set; }
            public int Quantity { get; set; }
            public string Unit { get; set; }
            public double PricePerUnit { get; set; }
        }
    }
}
