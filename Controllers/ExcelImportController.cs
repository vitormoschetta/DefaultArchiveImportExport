using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Dapper;
using DefaultArchiveImportExport.Data;
using DefaultArchiveImportExport.Models;
using DefaultArchiveImportExport.Util;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

namespace DefaultArchiveImportExport.Controllers
{
    public class ExcelImportController : Controller
    {
        private readonly Contexto _context;
        private readonly IConfiguration _configuration;
        private readonly ImportaDados _importaDados;

        public ExcelImportController(Contexto context, IConfiguration configuration, ImportaDados importaDados)
        {
            _context = context;    
            _configuration = configuration;
            _importaDados = importaDados;
        }

        // Using Dapper se for preciso
        public string GetConnection()
        {
            var connection = _configuration.GetSection("ConnectionStrings").GetSection("DefaultConnection").Value;
            return connection;
        }

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


        

        
        public IActionResult SignalR() => View();

        [HttpPost]
        public IActionResult SignalR(IFormFile arquivo)
        {                        
            string destino = "C:/uploadweb/";

            if (!Directory.Exists(destino)) Directory.CreateDirectory(destino);
            
            var fileName = destino +  System.IO.Path.GetFileName(arquivo.FileName);

            if (System.IO.File.Exists(fileName)) System.IO.File.Delete(fileName);
            
            using (var localFile = System.IO.File.OpenWrite(fileName))            
            using (var uploadedFile = arquivo.OpenReadStream())
            {
                uploadedFile.CopyTo(localFile);
                destino = localFile.Name;              
            }            
            
            _importaDados.Importar(destino);

            return View();
        }
      
    }
}