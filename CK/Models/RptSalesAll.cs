using System;
using System.Collections.Generic;

namespace CK.Models;

public partial class RptSalesAll
{
    public string? DpId { get; set; }

    public string? DpName { get; set; }

    public int? StoreIdR { get; set; }

    public string? StoreName { get; set; }

    public string? StoreFranchise { get; set; }

    public string ItemLookupCode { get; set; } = null!;

    public string? ItemName { get; set; }

    public DateTime? TransTime { get; set; }

    public DateTime? Yesterday { get; set; }

    public DateTime? TransDate { get; set; }

    public double? Qty { get; set; }

    public decimal Price { get; set; }

    public double? TotalSales { get; set; }

    public string TransactionNumber { get; set; }

    public decimal Cost { get; set; }

    public int? ByMonth { get; set; }

    public int? ByYear { get; set; }

    public int? ByDay { get; set; }

    public string StoreIdD { get; set; }

    public decimal Tax { get; set; }

    public double? TotalSalesTax { get; set; }

    public double? TotalSalesWithoutTax { get; set; }

    public double? TotalCostQty { get; set; }
    public int? SupplierId { get; set; }

    public string? SupplierCode { get; set; }

    public string? SupplierName { get; set; }
}
