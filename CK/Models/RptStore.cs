using System;
using System.Collections.Generic;

namespace CK.Models;

public partial class RptStore
{
    public string? StoreFranchise { get; set; }

    public int StoreId { get; set; }

    public string? StoreName { get; set; }

    public string? DpId { get; set; }

    public string? DpName { get; set; }

    public string ItemLookupCode { get; set; } = null!;

    public string? ItemName { get; set; }

    public decimal Cost { get; set; }

    public double Qty { get; set; }

    public string? SupplierCode { get; set; }

    public string? SupplierName { get; set; }

    public string? Dmanager { get; set; }

    public string? Fmanager { get; set; }
}
