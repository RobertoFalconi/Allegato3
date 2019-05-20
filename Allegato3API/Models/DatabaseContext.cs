using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace Allegato3API.Models
{
    public class DatabaseContext : DbContext
    {
        public DbSet<Templates> Template { get; set; }
        public DbSet<Files> File { get; set; }


        public DatabaseContext() : base("DefaultConnection")
        {

        }

        public System.Data.Entity.DbSet<Allegato3API.Models.Files> Files { get; set; }
    }
}
