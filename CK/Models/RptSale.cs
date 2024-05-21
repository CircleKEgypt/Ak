using System;
using System.Collections.Generic;

namespace CK.Models;

public partial class RptSale
{
    public string? DpId { get; set; }

    public int? GroupId { get; set; }

    public string? DpName { get; set; }

    public int? StoreCode { get; set; }

    public int StoreId { get; set; }

    public string? StoreName { get; set; }

    public string? StoreFranchise { get; set; }

    public int ItemId { get; set; }

    public string? ItemName { get; set; }

    public string? ItemLookupCode { get; set; }

    public DateTime? TransTime { get; set; }

    public int? ByDay { get; set; }

    public int? ByMonth { get; set; }

    public int? ByYear { get; set; }

    public DateTime? TransDate { get; set; }

    public double Qty { get; set; }

    public decimal Price { get; set; }

    public double? TotalSales { get; set; }

    public string? TransactionNumber { get; set; }

    public decimal Cost { get; set; }

    public double? TotalCostQty { get; set; }

    public double? Profit { get; set; }

    public decimal Tax { get; set; }

    public double? TotalSalesTax { get; set; }

    public double? TotalSalesWithoutTax { get; set; }

    public double? TotalCostWithoutTax { get; set; }

    public string? SupplierCode1 { get; set; }

    public string? Dmanager { get; set; }

    public string? Username { get; set; }

    public string? Fmanager { get; set; }
}
