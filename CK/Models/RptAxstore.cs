using System;
using System.Collections.Generic;

namespace CK.Models;

public partial class RptAxstore
{
    public string? StoreFranchise { get; set; }

    public string? StoreId { get; set; }

    public string? StoreName { get; set; }

    public DateTime Modified { get; set; }

    public string StoreNameInDy { get; set; } = null!;

    public decimal Qty { get; set; }

    public decimal? Cost { get; set; }

    public string ItemLookupCode { get; set; } = null!;

    public string? ItemName { get; set; }

    public string? DpId { get; set; }

    public string? DpName { get; set; }

    public string? SupplierCode { get; set; }

    public string? SupplierName { get; set; }

    public string? Dmanager { get; set; }

    public string? Fmanager { get; set; }

    public string? Username { get; set; }
}
