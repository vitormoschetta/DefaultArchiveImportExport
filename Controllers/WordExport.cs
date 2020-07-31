using System.IO;
using DefaultArchiveImportExport.Models;
using DefaultArchiveImportExport.Util;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;

namespace DefaultArchiveImportExport.Controllers
{
    public class WordExport : Controller
    {

        public IActionResult Index() => View();

        [HttpPost]
        public IActionResult Create(Processo modelo)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(mem, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
                {
                    wordDoc.AddMainDocumentPart();
                    // siga a ordem
                    Document doc = new Document();
                    Body body = new Body();

                    // 1 paragrafo
                    Paragraph para = new Paragraph();

                    ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                    ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Normal" };
                    Justification justification1 = new Justification() { Val = JustificationValues.Center };
                    ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();

                    paragraphProperties1.Append(paragraphStyleId1);
                    paragraphProperties1.Append(justification1);
                    paragraphProperties1.Append(paragraphMarkRunProperties1);

                    Run run = new Run();
                    RunProperties runProperties1 = new RunProperties();
                    
                    Text text = new Text() { Text = "EXCELENTÍSSIMO SENHOR DOUTOR JUIZ DE DIREITO DA MMª " + modelo.Vara + " DA COMARCA DE " + modelo.Comarca };

                    // siga a ordem 
                    run.Append(runProperties1);
                    run.Append(text);
                    para.Append(paragraphProperties1);
                    para.Append(run);

                    // 2 paragrafo
                    Paragraph para2 = new Paragraph();

                    ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                    ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Normal" };
                    Justification justification2 = new Justification() { Val = JustificationValues.Start };
                    ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();

                    paragraphProperties2.Append(paragraphStyleId2);
                    paragraphProperties2.Append(justification2);
                    paragraphProperties2.Append(paragraphMarkRunProperties2);

                    Run run2 = new Run();
                    RunProperties runProperties3 = new RunProperties();

                    run2.AppendChild(new Break());
                    run2.AppendChild(new Text("PROC. PROCESSO. " + modelo.NrProcessoCnj));
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Text(modelo.BancoNome + ", por seus advogados, nos autos da " + modelo.Acao +" que move em face de " + modelo.ClienteNome  + 
                                                " em trâmite nesta MMª Vara Cível, processo supra, em atenção ao R. Despacho de fls.  , respeitosamente vem expor e requerer o quanto segue:"));
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Text("01. A composição amigável entre as partes foi realizada por meio da Agência Bancária do Executado, o que ocasionou a não formalização de termo " +
                                                " de acordo para que o Autor possa juntar aos autos. "));
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Text("02. Cumpre informar as condições do acordo firmado entre as partes, por meio do qual o Requerido efetuará o pagamento da dívida pela importância de R$ " + 
                                                modelo.TotalDivida.ToString("C") + "(" + ConverteParaExtenso.ValorParaExtenso2(modelo.TotalDivida) + ") sendo entrada de R$ " + modelo.ValorEntrada.ToString("C") + 
                                                "(" + ConverteParaExtenso.ValorParaExtenso2(modelo.ValorEntrada) + ") paga em " +  modelo.DataVencimento.ToString("dd/MM/yyyy") + " mais " + 
                                                modelo.NrParcelas + "(" + ConverteParaExtenso.NumeroParaExtenso(modelo.NrParcelas) + ") parcelas de R$ " + modelo.ValorParcela.ToString("C")  + "(" + 
                                                ConverteParaExtenso.ValorParaExtenso2(modelo.ValorParcela) + ") nos meses subsequentes, sendo certo que não há intenção de renovação da dívida ante " +
                                                " a ausência de manifestação nesse sentido."));
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Text("03. Assim, requer a SUSPENSÃO DA AÇÃO PELO PRAZO CONCEDIDO AO REQUERIDO PARA QUE CUMPRA A OBRIGAÇÃO PACTUADA, nos termos, do art. 792, do Código de Processo Civil.  "));
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Break());
                    run2.AppendChild(new Text("04. Por fim, tendo em vista o convênio existente entre o SERASA e o Poder Judiciário, requer a expedição de ofício àquele órgão para exclusão do " +
                                                " nome do Requerido, efetivado em razão do ajuizamento da presente ação. "));

                    para2.Append(paragraphProperties2);
                    para2.Append(run2);

                    // todos os 2 paragrafos no main body
                    body.Append(para);
                    body.Append(para2);

                    doc.Append(body);

                    wordDoc.MainDocumentPart.Document = doc;

                    wordDoc.Close();
                }

                return File(mem.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "ABC.docx");
            }
        }
    }
}