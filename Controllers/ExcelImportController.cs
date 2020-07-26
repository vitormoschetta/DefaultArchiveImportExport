using System;
using System.Collections.Generic;
using System.IO;
using DefaultArchiveImportExport.Models;
using DefaultArchiveImportExport.Util;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace DefaultArchiveImportExport.Controllers
{
    public class ExcelImportController : Controller
    {
        public IActionResult Index() => View();

        public IActionResult Import()
        {                        
            var filePath = FileInputUtil.GetFileInfo("Data", "ExcelProducts.xlsx").FullName;			

			using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
			{                
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];  // => Pega o primeiro arquivo com o nome "ExcelProducts"
                     
                var rowCount = worksheet.Dimension.End.Row; // => Identifica quantas linhas preenchidas tem o arquivo   

                //var colCnt = worksheet.Dimension.End.Column + 1; // => Identifica quantas colunas preenchidas tem o arquivo

                IList<Product> products = new List<Product>(); 

				for (int row = 2; row <= rowCount; row++) 
                {       
                    Product product = new Product();
                    for (int col = 1; col < 4; col++) 
                    {
                        if (col == 1) product.Id = new Guid(worksheet.Cells[row, col].Value.ToString());                 
                        if (col == 2) product.Name = worksheet.Cells[row, col].Value.ToString();
                        if (col == 3) product.Price = Convert.ToDecimal(worksheet.Cells[row, col].Value);
                    }
                    products.Add(product);
                }

                return View(products);                
            }
        }
    }
}