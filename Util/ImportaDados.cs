using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using DefaultArchiveImportExport.Data;
using DefaultArchiveImportExport.Models;
using Microsoft.AspNetCore.SignalR;
using OfficeOpenXml;

namespace DefaultArchiveImportExport.Util
{
    public class ImportaDados: Hub
    {
        private readonly IHubContext<ImportaDados> _streaming;
        private readonly Contexto _context;
        public ImportaDados(Contexto context, IHubContext<ImportaDados> streaming)
        {
            _context = context;
            _streaming = streaming;
        }

        private async Task WriteOnStream(string Mensagem,string total, string atual)
        {
            string jsonData = string.Format("{0}\n", JsonSerializer.Serialize(new { Mensagem, total, atual }));
            await _streaming.Clients.All.SendAsync("ReceiveMessage", jsonData);
        }

        public async void Importar(string caminhoArquivo)
        {
            //var filePath = FileInputUtil.GetFileInfo("Data/TestImportExcel",  Arquivo).FullName; // => Diretório e nome do arquivo		

            using (ExcelPackage package = new ExcelPackage(new FileInfo(caminhoArquivo)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];  // => Pega o primeiro arquivo com o nome "ExcelProducts"

                int rowCount = worksheet.Dimension.End.Row; // => Identifica quantas linhas preenchidas tem o arquivo   

                var colCount = worksheet.Dimension.End.Column; // => Identifica quantas colunas preenchidas tem o arquivo

                await ImportarDados(worksheet, rowCount, colCount);
                
            }
        }

        public async Task ImportarDados(ExcelWorksheet worksheet, int rowCount, int colCount)
        {                                    
            int contador = 1;
         
            for (int row = 2; row <= rowCount; row++) 
            {       
                Product product = new Product();
                for (int col = 1; col <= colCount; col++) 
                {
                    if (col == 1) product.Id = Guid.NewGuid();                 
                    if (col == 2) product.Name = worksheet.Cells[row, col].Value.ToString();
                    if (col == 3) product.Price = Convert.ToDecimal(worksheet.Cells[row, col].Value);
                }
                
                _context.Add(product);
                _context.SaveChanges();
                
                await WriteOnStream("Registro.: " + Convert.ToString(contador)  + " de " + Convert.ToString(rowCount), rowCount.ToString(), Convert.ToString(contador));
                
                contador++;
            }          
            
        }
    }
}