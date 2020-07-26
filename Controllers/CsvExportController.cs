using System;
using System.Collections.Generic;
using System.Text;
using DefaultArchiveImportExport.Data;
using DefaultArchiveImportExport.Models;
using Microsoft.AspNetCore.Mvc;

namespace DefaultArchiveImportExport.Controllers
{
    public class CsvExportController: Controller
    {
        public IActionResult Index() => View();
        public IActionResult InputToCsv() => View();
        public IActionResult DataToCsv() => View();



        public IActionResult ExportData()
        {            
            var products = DataLoader.GetProducts();
            return ExportExcel(products);
        }

        

        [HttpPost]
        public IActionResult ExportInput(Product product)
        {            
            product.Id = Guid.NewGuid();

            IList<Product> products = new List<Product>();            
            products.Add(product);

            return ExportExcel(products);
        }


        
        public IActionResult ExportExcel(IEnumerable<Product> products)
        {
            var builder = new StringBuilder();
            builder.AppendLine("Id,Customer,Date,Total");

            foreach (var item in products)
            {
                builder.AppendLine($"{item.Id},{item.Name},{item.Price}");
            }

            var nomeArquivo = "CsvProducts";

            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", nomeArquivo + ".csv");
            
        }
    }
}