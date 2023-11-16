using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace ExcelReader
{
    internal class CitizenDbContext: DbContext
    {
        public DbSet<Citizen> Citizens { get; set; }

        private static string ConnectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // Replace "YourConnectionString" with the connection string for your SQL Server

            optionsBuilder.UseSqlServer(ConnectionString);
        }

        public List<Citizen> GetAllCitizens()
        {
            return Citizens.ToList();
        }
    }
}
