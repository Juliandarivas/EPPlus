using System;
using System.Collections.Generic;
using System.IO;
using EPPlus.Model.Models;
using EPPlus.Model.Properties;
using EPPlus.v4._5._3._2;

namespace EPPlus.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Product> products = CreateProducts();

            byte[] bytes = new EPPlusGenerator().Generate(Resources.Plantilla, products);

            File.WriteAllBytes(@"C:\temp\Products.xlsx", bytes);

            System.Console.WriteLine("Generado en C:\\temp\\Products.xlsx");
        }

        private static List<Product> CreateProducts()
        {
            return new List<Product>
            {
                new Product(1, "FNX0001", "Silla Oficina Paris", new DateTime(2020, 5, 1), 500_000),
                new Product(2, "FNX0002", "Silla Oficina Dublin", new DateTime(2020, 5, 1), 3_000_000),
                new Product(3, "FNX0003", "Silla Oficina Roma", new DateTime(2020, 5, 1), 400_000),
                new Product(4, "FNX0004", "Silla Oficina Berlin", new DateTime(2020, 5, 1), 800_000),
                new Product(5, "FNX0005", "Sillón El Cairo", new DateTime(2020, 5, 1), 1_200_000),
                new Product(6, "FNX0006", "Escritorio sala de juntas New York", new DateTime(1980, 1, 1), 1_500_000),
                new Product(7, "FNX0007", "Silla rimax", new DateTime(2020, 5, 1), 2_000_000),
                new Product(8, "FNX0008", "Alfombra sala", new DateTime(2020, 5, 1), 3_100_000),
                new Product(9, "FNX0009", "Mesa para comedor de vidrio", new DateTime(2020, 5, 1), 2_000_000),
                new Product(10, "FNX0010", "Mesa redonda de vidrio", new DateTime(2020, 5, 1), 5_000_000),
            };
        }
    }
}
