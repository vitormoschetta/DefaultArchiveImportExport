# DefaultArchiveImportExport

Um projeto para importação e exportação de Arquivos Word, Excel, CSV, etc..


#### Adicionar SignalR para acompanhamento do processo de importação em tempo real:

###### Add Classe ImportaDados / Util:

```
public class ImportaDados: Hub
{
    private readonly IHubContext<ImportaDados> _streaming;
    private readonly Contexto _context;
    public ImportaDados(Contexto context, IHubContext<ImportaDados> streaming)
    {
        _context = context;
        _streaming = streaming;
    }
    
    // Método que aciona o túnel/conversa contínua Servidor-Cliente:
    private async Task WriteOnStream(string Mensagem,string total, string atual)
    {
        string jsonData = string.Format("{0}\n", JsonSerializer.Serialize(new { Mensagem, total, atual }));
        await _streaming.Clients.All.SendAsync("ReceiveMessage", jsonData);
    }    
}

```

Nessa mesma classe (ImportaDados), será adicionado 
