using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Text;

namespace EntityFrameworkCoreSample.Model
{
    public class VehicleStoreContext : DbContext
    {
        public DbSet<BranchEntity> Branches { get; set; }

        public DbSet<MonthlyRevenueEntity> Revenues { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlite("Data Source=vehicle_store.db");
        }
    }
}
