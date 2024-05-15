using System;
using System.Collections.Generic;

namespace CK.Models;

public partial class RptSalesAll
{
    public string? DpId { get; set; }

    public string? DpName { get; set; }

    public string? StoreId { get; set; }

    public string? StoreIdR { get; set; }

    public string? StoreName { get; set; }

    public string? StoreFranchise { get; set; }

    public string? ItemLookupCode { get; set; }

    public string? ItemName { get; set; }

    public DateTime? TransTime { get; set; }

    public DateOnly? TransDate { get; set; }

    public double? Qty { get; set; }

    public double? Price { get; set; }

    public double? TotalSales { get; set; }

    public string? TransactionNumber { get; set; }

    public double? Cost { get; set; }

    public int? ByMonth { get; set; }

    public int? ByYear { get; set; }

    public int? ByDay { get; set; }

    public int? StoreIdD { get; set; }

    public double? Tax { get; set; }

    public double? TotalSalesTax { get; set; }

    public double? TotalSalesWithoutTax { get; set; }

    public double? TotalCostQty { get; set; }

    public string? Dmanager { get; set; }

    public string? Fmanager { get; set; }

    public string? Username { get; set; }
}
