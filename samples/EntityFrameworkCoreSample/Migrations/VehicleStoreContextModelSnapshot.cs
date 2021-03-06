﻿// <auto-generated />
using System;
using EntityFrameworkCoreSample.Model;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

namespace EntityFrameworkCoreSample.Migrations
{
    [DbContext(typeof(VehicleStoreContext))]
    partial class VehicleStoreContextModelSnapshot : ModelSnapshot
    {
        protected override void BuildModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "2.2.6-servicing-10079");

            modelBuilder.Entity("EntityFrameworkCoreSample.Model.BranchEntity", b =>
                {
                    b.Property<string>("Id")
                        .ValueGeneratedOnAdd();

                    b.Property<string>("Location");

                    b.Property<string>("Name");

                    b.Property<string>("Phone");

                    b.HasKey("Id");

                    b.ToTable("Branches");
                });

            modelBuilder.Entity("EntityFrameworkCoreSample.Model.MonthlyRevenueEntity", b =>
                {
                    b.Property<Guid>("Id")
                        .ValueGeneratedOnAdd();

                    b.Property<string>("BranchId");

                    b.Property<DateTime>("MonthYear");

                    b.Property<decimal>("Value");

                    b.HasKey("Id");

                    b.HasIndex("BranchId");

                    b.ToTable("Revenues");
                });

            modelBuilder.Entity("EntityFrameworkCoreSample.Model.MonthlyRevenueEntity", b =>
                {
                    b.HasOne("EntityFrameworkCoreSample.Model.BranchEntity", "Branch")
                        .WithMany("Revenues")
                        .HasForeignKey("BranchId");
                });
#pragma warning restore 612, 618
        }
    }
}
