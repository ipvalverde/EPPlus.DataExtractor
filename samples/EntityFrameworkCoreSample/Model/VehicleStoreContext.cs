using Microsoft.EntityFrameworkCore;

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

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<BranchEntity>()
                .ToTable(BranchEntity.TableName)
                .HasMany(b => b.Revenues)
                .WithOne(r => r.Branch);

            modelBuilder.Entity<MonthlyRevenueEntity>()
                .ToTable(MonthlyRevenueEntity.TableName);
        }
    }
}
