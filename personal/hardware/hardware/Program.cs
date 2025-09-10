using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace hardware
{
    class Program
    {
        static void Main(string[] args)
        {
            List<ComputerHardware> computerHardware = new List<ComputerHardware>() {
                new ComputerHardware() { Name = "Intel Core i7-9700K", Type = "CPU", Price = 400, ClockSpeed = 3.6, Cores = 8, Brand = "Intel" },
                new ComputerHardware() { Name = "AMD Ryzen 9 5950X", Type = "CPU", Price = 700, ClockSpeed = 3.4, Cores = 16, Brand = "AMD" },
                new ComputerHardware() { Name = "NVIDIA GeForce RTX 3080", Type = "GPU", Price = 700, ClockSpeed = 1.7, Cores = 8704, Brand = "NVIDIA" },
                new ComputerHardware() { Name = "AMD Radeon RX 6800 XT", Type = "GPU", Price = 650, ClockSpeed = 2.0, Cores = 72, Brand = "AMD" },
                new ComputerHardware() { Name = "Intel Core i5-10400", Type = "CPU", Price = 200, ClockSpeed = 2.9, Cores = 6, Brand = "Intel" },
                new ComputerHardware() { Name = "AMD Ryzen 5 5600X", Type = "CPU", Price = 300, ClockSpeed = 3.7, Cores = 6, Brand = "AMD" },
                new ComputerHardware() { Name = "NVIDIA GeForce RTX 3060 Ti", Type = "GPU", Price = 400, ClockSpeed = 1.6, Cores = 4864, Brand = "NVIDIA" },
                new ComputerHardware() { Name = "AMD Radeon RX 6700 XT", Type = "GPU", Price = 400, ClockSpeed = 2.4, Cores = 40, Brand = "AMD" },
                new ComputerHardware() { Name = "Intel Core i9-11900K", Type = "CPU", Price = 500, ClockSpeed = 3.2, Cores = 10, Brand = "Intel" },
                new ComputerHardware() { Name = "AMD Ryzen 7 5800X", Type = "CPU", Price = 350, ClockSpeed = 3.9, Cores = 8, Brand = "AMD" },
                new ComputerHardware() { Name = "NVIDIA GeForce RTX 3090", Type = "GPU", Price = 1500, ClockSpeed = 1.4, Cores = 10496, Brand = "NVIDIA" },
                new ComputerHardware() { Name = "AMD Radeon RX 6900 XT", Type = "GPU", Price = 1000, ClockSpeed = 2.0, Cores = 80, Brand = "AMD" },
                new ComputerHardware() { Name = "Intel Core i3-10100", Type = "CPU", Price = 150, ClockSpeed = 3.6, Cores = 4, Brand = "Intel" },
                new ComputerHardware() { Name = "AMD Ryzen 3 5600X", Type = "CPU", Price = 250, ClockSpeed = 3.6, Cores = 6, Brand = "AMD" },
                new ComputerHardware() { Name = "NVIDIA GeForce RTX 3070", Type = "GPU", Price = 500, ClockSpeed = 1.5, Cores = 5888, Brand = "NVIDIA" },
                new ComputerHardware() { Name = "AMD Radeon RX 6700", Type = "GPU", Price = 350, ClockSpeed = 2.3, Cores = 36, Brand = "AMD" },
                new ComputerHardware() { Name = "Intel Core i9-9900K", Type = "CPU", Price = 450, ClockSpeed = 3.2, Cores = 8, Brand = "Intel" },
                new ComputerHardware() { Name = "AMD Ryzen 7 3700X", Type = "CPU", Price = 300, ClockSpeed = 3.6, Cores = 8, Brand = "AMD" },
                new ComputerHardware() { Name = "NVIDIA GeForce RTX 3080 Ti", Type = "GPU", Price = 1200, ClockSpeed = 1.6, Cores = 5888, Brand = "NVIDIA" },
                new ComputerHardware() { Name = "AMD Radeon RX 6800", Type = "GPU", Price = 600, ClockSpeed = 1.8, Cores = 64, Brand = "AMD" }
            };

            Console.WriteLine("1. Pas CPU\n2. Prix > X\n3. CPUs mauvais pour jouer\n4. Configs potables\n5. Configs AMD");
            int choix = int.Parse(Console.ReadLine());

            IEnumerable<ComputerHardware> query = choix switch
            {
                1 => computerHardware.Where(h => h.Type != "CPU"),
                2 => computerHardware.Where(h => h.Price > AskNumber("Entrez le prix limite : ")),
                3 => computerHardware.Where(h => h.Type == "CPU" && (h.ClockSpeed < 3.0 || h.Cores < 4)),
                4 => computerHardware.Where(h => (h.Type == "GPU" && h.Cores >= 32) || (h.Type == "CPU" && h.Cores >= 8)),
                5 => computerHardware.Where(h => h.Brand == "AMD"),
                _ => computerHardware
            };

            Console.WriteLine("Trier par (1) prix ou (2) horloge ?");
            int critere = int.Parse(Console.ReadLine());
            Console.WriteLine("Ordre (1) croissant ou (2) décroissant ?");
            int ordre = int.Parse(Console.ReadLine());

            //tuple
            query = (critere, ordre) switch
            {
                (1, 1) => query.OrderBy(h => h.Price),
                (1, 2) => query.OrderByDescending(h => h.Price),
                (2, 1) => query.OrderBy(h => h.ClockSpeed),
                (2, 2) => query.OrderByDescending(h => h.ClockSpeed),
                _ => query //default
            };

            //ChatGPT did this part
            /*   foreach (var h in query)
                    Console.WriteLine($"{h.Name} ({h.Type}) - {h.Price} CHF - {h.ClockSpeed} GHz - {h.Cores} cœurs - {h.Brand}");

                Console.WriteLine("Exporter en CSV ? (o/n)");
                string reponse = Console.ReadLine();

                if (reponse?.ToLower() == "o")
                {
                    File.WriteAllLines("export.csv",
                        new[] { "Name;Type;Price;ClockSpeed;Cores;Brand" }
                        .Concat(query.Select(h => $"{h.Name};{h.Type};{h.Price};{h.ClockSpeed};{h.Cores};{h.Brand}")));

                    Console.WriteLine("Export terminé -> export.csv");
                }
                else
                {
                    Console.WriteLine("No worries, press any key...");
                    Console.ReadKey();
                }*/
        }

        static double AskNumber(string message)
        {
            Console.Write(message);
            return double.Parse(Console.ReadLine());
        }
    }

    class ComputerHardware
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public double Price { get; set; }
        public double ClockSpeed { get; set; }
        public int Cores { get; set; }
        public string Brand { get; set; }
    }
}
