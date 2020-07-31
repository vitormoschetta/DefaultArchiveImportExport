using System;

namespace DefaultArchiveImportExport.Models
{
    public class Processo
    {
        public string Vara { get; set; }
        public string Comarca { get; set; }
        public string NrProcessoCnj { get; set; }        
        public string Acao { get; set; }
        public string BancoNome { get; set; }
        public string ClienteNome { get; set; }
        public decimal TotalDivida { get; set; }
        public decimal ValorEntrada { get; set; }
        public DateTime DataVencimento { get; set; }
        public int NrParcelas { get; set; }
        public decimal ValorParcela { get; set; } 
    }
}