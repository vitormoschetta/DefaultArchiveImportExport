using System;
using System.Collections.Generic;
using DefaultArchiveImportExport.Models;

namespace DefaultArchiveImportExport.Data
{
    public static class DataLoader
    {
        public static IEnumerable<Product> GetProducts()
        {
            return new[]{
                new Product{ Id = Guid.NewGuid(), Name = "Product01", Price = 2.99M  },
                new Product{ Id = Guid.NewGuid(), Name = "Product02", Price = 5.50M },
                new Product{ Id = Guid.NewGuid(), Name = "Product03", Price = 10.00M },
                new Product{ Id = Guid.NewGuid(), Name = "Product04", Price = 3.50M },
                new Product{ Id = Guid.NewGuid(), Name = "Product05", Price = 100.00M },
            };
        }
    }
}