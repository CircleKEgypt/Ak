using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;
using System.ComponentModel.DataAnnotations.Schema;

namespace CK.Models
{

    public partial class ServerConnection : DbContext
    {
        public ServerConnection(string databaseConnection)
    : base()
        {
            ConnectionString = databaseConnection;
        }


        public string ConnectionString { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(ConnectionString,

             mySqlOptions =>
     mySqlOptions.CommandTimeout(220).EnableRetryOnFailure(
         maxRetryCount: 10,
         maxRetryDelay: TimeSpan.FromSeconds(220),
         errorNumbersToAdd: null)

         );
        }

        //public DbSet<RetailtransactionsalestransST>? Retailtransactionsalestrans { get; set; }

        //public DbSet<RetailtransactiontableST>? Retailtransactiontable { get; set; }
        //public DbSet<RetailtransactionpaymenttransST>? Retailtransactionpaymenttrans { get; set; }
        //public DbSet<RetailtransactionpaymenttransCO>? RetailtransactionpaymenttransC { get; set; }
        //public DbSet<Inventtablemodule>? Inventtablemodule { get; set; }


    }




}
