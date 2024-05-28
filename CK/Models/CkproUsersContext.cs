using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace CK.Models;

public partial class CkproUsersContext : DbContext
{
    public CkproUsersContext()
    {
    }

    public CkproUsersContext(DbContextOptions<CkproUsersContext> options)
        : base(options)
    {
    }

    public virtual DbSet<RptUser> RptUsers { get; set; }

    public virtual DbSet<RptUsers2> RptUsers2s { get; set; }

    public virtual DbSet<Storeuser> Storeusers { get; set; }

    public virtual DbSet<User> Users { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see https://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=192.168.1.156;User ID=sa;Password=P@ssw0rd;Database=CkproUsers;Connect Timeout=150;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<RptUser>(entity =>
        {
            entity
                .HasNoKey()
                .ToView("RptUsers");

            entity.Property(e => e.Department).HasMaxLength(100);
            entity.Property(e => e.Dmanager)
                .HasMaxLength(255)
                .HasColumnName("DManager");
            entity.Property(e => e.Fmanager)
                .HasMaxLength(255)
                .HasColumnName("FManager");
            entity.Property(e => e.Inventlocation).HasMaxLength(255);
            entity.Property(e => e.Password).HasMaxLength(255);
            entity.Property(e => e.RmsstoNumber)
                .HasMaxLength(255)
                .HasColumnName("RMSstoNumber");
            entity.Property(e => e.Role).HasMaxLength(100);
            entity.Property(e => e.Server).HasMaxLength(255);
            entity.Property(e => e.Storenumber)
                .HasMaxLength(255)
                .HasColumnName("storenumber");
            entity.Property(e => e.Username).HasMaxLength(255);
            entity.Property(e => e.Username2).HasMaxLength(100);
        });

        modelBuilder.Entity<RptUsers2>(entity =>
        {
            entity
                .HasNoKey()
                .ToView("RptUsers2");

            entity.Property(e => e.Department).HasMaxLength(100);
            entity.Property(e => e.Dmanager)
                .HasMaxLength(255)
                .HasColumnName("DManager");
            entity.Property(e => e.Fmanager)
                .HasMaxLength(255)
                .HasColumnName("FManager");
            entity.Property(e => e.Password).HasMaxLength(255);
            entity.Property(e => e.Role).HasMaxLength(100);
            entity.Property(e => e.Server).HasMaxLength(255);
            entity.Property(e => e.Storenumber)
                .HasMaxLength(255)
                .HasColumnName("storenumber");
            entity.Property(e => e.Username).HasMaxLength(255);
        });

        modelBuilder.Entity<Storeuser>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK_st");

            entity.ToTable("Storeuser");

            entity.HasIndex(e => new { e.Inventlocation, e.Storenumber, e.Username, e.Password, e.Name, e.Server, e.RmsstoNumber }, "index_name");

            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.ArabicN).HasMaxLength(255);
            entity.Property(e => e.Company).HasMaxLength(255);
            entity.Property(e => e.Dbase).HasMaxLength(255);
            entity.Property(e => e.District).HasMaxLength(255);
            entity.Property(e => e.Dmanager)
                .HasMaxLength(255)
                .HasColumnName("DManager");
            entity.Property(e => e.Email).HasMaxLength(255);
            entity.Property(e => e.Fmanager)
                .HasMaxLength(255)
                .HasColumnName("FManager");
            entity.Property(e => e.Franchise).HasMaxLength(50);
            entity.Property(e => e.Inventlocation).HasMaxLength(255);
            entity.Property(e => e.Name).HasMaxLength(255);
            entity.Property(e => e.Password).HasMaxLength(255);
            entity.Property(e => e.PriceCategory).HasMaxLength(255);
            entity.Property(e => e.RmsstoNumber)
                .HasMaxLength(255)
                .HasColumnName("RMSstoNumber");
            entity.Property(e => e.Server).HasMaxLength(255);
            entity.Property(e => e.StartDate).HasMaxLength(50);
            entity.Property(e => e.Storenumber)
                .HasMaxLength(255)
                .HasColumnName("storenumber");
            entity.Property(e => e.Username).HasMaxLength(255);
            entity.Property(e => e.Zkip)
                .HasMaxLength(255)
                .HasColumnName("ZKIP");
        });

        modelBuilder.Entity<User>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__Users__3214EC27018A4CE6");

            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.Department).HasMaxLength(100);
            entity.Property(e => e.Password).HasMaxLength(100);
            entity.Property(e => e.Role).HasMaxLength(100);
            entity.Property(e => e.User1)
                .HasMaxLength(100)
                .HasColumnName("User");
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
