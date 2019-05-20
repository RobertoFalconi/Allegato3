using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace Allegato3API.Models
{
    public class Allegato3APIContext : DbContext
    {

        public DbSet<Templates> Templates { get; set; }
        public DbSet<Files> Files { get; set; }

        public Allegato3APIContext() : base("name=DefaultConnection")
        {
        }

        //public System.Data.Entity.DbSet<Allegato3API.Models.Files> Files { get; set; }

        //protected override void OnModelCreating(DbModelBuilder modelBuilder)
        //{
        //    base.OnModelCreating(modelBuilder);
        //    modelBuilder.Entity<Files>().ToTable("db", "Files");
        //}

    }
}
