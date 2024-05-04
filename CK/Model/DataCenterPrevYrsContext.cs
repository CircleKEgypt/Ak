using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace CK.Model;

public partial class DataCenterPrevYrsContext : DbContext
{
    public DataCenterPrevYrsContext()
    {
    }

    public DataCenterPrevYrsContext(DbContextOptions<DataCenterPrevYrsContext> options)
        : base(options)
    {
    }

    public virtual DbSet<RptSales2> RptSales2s { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see https://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=192.168.1.156;User ID=sa;Password=P@ssw0rd;Database=DATA_CENTER_Prev_Yrs;Connect Timeout=150;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<RptSales2>(entity =>
        {
            entity
                .HasNoKey()
                .ToView("RptSales2");

            entity.Property(e => e.Cost).HasColumnType("money");
            entity.Property(e => e.DpId)
                .HasMaxLength(17)
                .HasColumnName("DpID");
            entity.Property(e => e.DpName).HasMaxLength(30);
            entity.Property(e => e.ItemId).HasColumnName("ItemID");
            entity.Property(e => e.ItemLookupCode).HasMaxLength(25);
            entity.Property(e => e.ItemName).HasMaxLength(30);
            entity.Property(e => e.Price).HasColumnType("money");
            entity.Property(e => e.StoreFirstTransactionDate).HasColumnName("Store_FIRST_TRANSACTION_DATE");
            entity.Property(e => e.StoreFranchise)
                .HasMaxLength(50)
                .HasColumnName("StoreFranchise");
            entity.Property(e => e.StoreGov).HasMaxLength(50);
            entity.Property(e => e.StoreId).HasColumnName("StoreID");
            entity.Property(e => e.StoreName)
                .HasMaxLength(50)
                .HasColumnName("StoreName");
            entity.Property(e => e.StoreNameD365).HasMaxLength(50);
            entity.Property(e => e.StorePriceCat)
                .HasMaxLength(10)
                .IsFixedLength();
            entity.Property(e => e.StoreType).HasMaxLength(50);
            entity.Property(e => e.SupplierName).HasMaxLength(30);
            entity.Property(e => e.Tax).HasColumnType("money");
            entity.Property(e => e.TransTime).HasColumnType("datetime");
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
