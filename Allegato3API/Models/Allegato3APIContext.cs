using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace Allegato3API.Models
{
    public class Allegato3APIContext : DbContext
    {

        public DbSet<Files> Files { get; set; }

        public Allegato3APIContext() : base("name=DefaultConnection")
        {
        }

    }
}
