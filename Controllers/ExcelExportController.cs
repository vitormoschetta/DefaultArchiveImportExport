using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DefaultArchiveImportExport.Data;
using DefaultArchiveImportExport.Models;
using Microsoft.AspNetCore.Mvc;

namespace DefaultArchiveImportExport.Controllers
{
    public class ExcelExportController: Controller
    {
        public IActionResult Index() => View();
        public IActionResult InputToExcel() => View();
        public IActionResult DataToExcel() => View();



        public async Task<IActionResult> ExportData()
        {            
            var products = DataLoader.GetProducts();
            return await ExportExcel(products);
        }

        

        [HttpPost]
        public async Task<IActionResult> ExportInput(Product product)
        {            
            product.Id = Guid.NewGuid();

            IList<Product> products = new List<Product>();            
            products.Add(product);

            return await ExportExcel(products);
        }


        
        public async Task<IActionResult> ExportExcel(IEnumerable<Product> products)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("sms");
            var currentRow = 1;

            worksheet.Cell(currentRow, 1).Value = "Id";
            worksheet.Cell(currentRow, 2).Value = "Name";      
            worksheet.Cell(currentRow, 3).Value = "Price";                      

            foreach (var item in products)
            {
                currentRow++;
                worksheet.Cell(currentRow, 1).Value = item.Id;
                worksheet.Cell(currentRow, 2).Value = item.Name;
                worksheet.Cell(currentRow, 3).Value = item.Price;                
            }      

			// Auto-ajuste colunas
            worksheet.Columns().AdjustToContents();

            await using var memory = new MemoryStream();

            workbook.SaveAs(memory);        

            var nomeArquivo = "ExcelProducts";    

            return File(memory.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nomeArquivo + ".xlsx");
            
        }
    }

}