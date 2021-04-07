using EnvioCorreo_ListaSP.DTO;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace EnvioCorreo_ListaSP.DBContext
{
    public class InterfacesDBContext : DbContext
    {
        public InterfacesDBContext()
        {
        }

        public InterfacesDBContext(DbContextOptions<InterfacesDBContext> options)
           : base(options)
        {
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            var Conn = ConfigurationManager.ConnectionStrings["InterfacesDB"].ConnectionString;
            optionsBuilder.UseSqlServer(Conn);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<EROfisisDTO>().HasNoKey().ToView(null);
        }
    }
}
