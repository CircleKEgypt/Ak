namespace CK.Models
{
    public class SalesParameters
    {
        public string startDate { get; set; }
        public string endDate { get; set; }
        public string Store { get; set; }
        public List<string> SelectedStores { get; set; }
        public string Department { get; set; }
        public string Supplier { get; set; }
        public bool ExportAfterClick { get; set; }
        public string[] SelectedItems { get; set; }
        public bool VPerDay { get; set; }
        public bool VPerMonYear { get; set; }
        public bool VPerMon { get; set; }
        public bool VPerYear { get; set; }
        public bool VQty { get; set; }
        public bool VPrice { get; set; }
        public bool VStoreName { get; set; }
        public bool VDepartment { get; set; }
        public bool VTotalSales { get; set; }
        public bool VTotalCost { get; set; }
        public bool VTotalTax { get; set; }
        public bool VTotalSalesTax { get; set; }
        public bool VTotalSalesWithoutTax { get; set; }
        public bool VTotalCostQty { get; set; }
        public bool VCost { get; set; }
        public bool VItemLookupCode { get; set; }
        public bool VItemName { get; set; }
        public bool VSupplierId { get; set; }
        public bool VSupplierName { get; set; }
        public string Franchise { get; set; }
        public bool VTransactionNumber { get; set; }
        public bool VFranchise { get; set; }
        public int? MonthToFilter { get; set; }
        public string ItemLookupCodeTxt { get; set; }
        public string ItemNameTxt { get; set; }
        public bool TMT { get; set; }
        public bool RMS { get; set; }
        public bool DBbefore { get; set; }
        public bool Yesterday { get; set; }
        public bool VTransactionCount { get; set; }
        public bool exportAfterClick { get; set;}
        public bool VStoreId { get; set;}
        public bool VDpId { get; set; }
        public bool VPaidtype { get; set; }
        public bool VDateInTime { get; set; }
        public bool Vbatch { get;set; }
    }
}
//change suppliercode, transactionnumber,storeid,dpid to string in rptsales and 2 and all 