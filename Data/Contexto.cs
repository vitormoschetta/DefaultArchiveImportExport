using DefaultArchiveImportExport.Models;
using Microsoft.EntityFrameworkCore;

namespace DefaultArchiveImportExport.Data
{
    public class Contexto: DbContext
{
    public Contexto(DbContextOptions<Contexto> options) : base(options)
    {
    }
    public DbSet<Product> Product { get; set; }
}
}