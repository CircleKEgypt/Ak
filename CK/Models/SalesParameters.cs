using System.Security.Cryptography.X509Certificates;

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
        public bool Vbatch { get; set; }
        public bool StockFromBranch { get; set; }
        public string linkRmsRptStoreView = @"
								SELECT st.Franchise StoreFranchise,it.StoreID,sty.Username StoreName,
								Dep.code DpId, Dep.Name dpName,It.ItemLookupCode, It.Description ItemName,It.Cost--, It.Price 
								, It.Quantity Qty
								,Supp.Code SupplierCode ,Supp.SupplierName,sty.Username,sty.DManager,sty.FManager
								FROM [192.168.1.156].[DATA_CENTER].[dbo].[Item] AS It
								inner JOIN [192.168.1.156].[DATA_CENTER].[dbo].[department] AS Dep ON It.DepartmentID = Dep.ID AND It.storeid = Dep.storeid
								left JOIN [192.168.1.156].[DATA_CENTER].[dbo].[SupplierList] AS SuppL ON It.storeid = SuppL.storeid 
								AND It.SupplierID = SuppL.SupplierID AND It.ID = SuppL.ItemID
								left JOIN [192.168.1.156].[DATA_CENTER].[dbo].[Supplier] AS Supp ON SuppL.SupplierID = Supp.ID AND SuppL.storeid = Supp.storeid
								left join (select RMSstoNumber,Username,DManager,FManager from [192.168.1.156].CkproUsers.dbo.Storeuser) sty on sty.RMSstoNumber =convert(varchar(10),it.storeid)
								left join [192.168.1.156].[DATA_CENTER].dbo.STORES st on st.STORE_ID =it.storeid
								where st.Franchise='SUB-FRANCHISE' and sty.RMSstoNumber !='58'";
        public string linkDyRptStoreView = @"
select st.Franchise StoreFranchise,st.Storenumber StoreID,
st.username StoreName,
Inv.Modifieddatetime Modified, 
Inv.Wmslocationid StoreNameInDy, 
Inv.Physicalinvent Qty,t.COSTPRICE Cost, Inv.Itemid ItemLookupCode,
It.Name ItemName, CateN.CODE DpId,CateN.Name dpName
,ca.Primaryvendorid SupplierCode,W.SupplierName
,DManager,FManager,Username
 from [192.168.1.210].AXDB.dbo.Inventsum Inv
left join [192.168.1.210].AXDB.dbo.Inventtable ca on Inv.Itemid = ca.Itemid
left join [192.168.1.156].DATA_CENTER.dbo.supplier w on w.Code=ca.Primaryvendorid collate SQL_Latin1_General_CP1_CI_AS
left join [192.168.1.210].AXDB.dbo.Ecoresproducttranslation It on ca.Product = It.Product
left join [192.168.1.210].AXDB.dbo.Ecoresproductcategory Re on It.Product = Re.Product
left join [192.168.1.210].AXDB.dbo.Ecorescategory CateN on Re.Category = CateN.Recid
left join [192.168.1.156].CkproUsers.dbo.Storeuser st on st.Inventlocation=Inv.Wmslocationid
left join (SELECT distinct s.COSTPRICE,s.ITEMID a FROM  [192.168.1.210].AXDB.dbo.Salesline s WHERE s.Confirmeddlv = (
        SELECT  MAX(Confirmeddlv) FROM  [192.168.1.210].AXDB.dbo.Salesline where itemid=s.ITEMID) )t on t.a=inv.ITEMID
where ca.Dataareaid = 'tmt'";
        public string RmsRptStoreView = @"
								SELECT st.Franchise StoreFranchise,it.StoreID,sty.Username StoreName,
								Dep.code DpId, Dep.Name dpName,It.ItemLookupCode, It.Description ItemName,It.Cost--, It.Price 
								, It.Quantity Qty
								,Supp.Code SupplierCode ,Supp.SupplierName,sty.Username,sty.DManager,sty.FManager
								FROM [DATA_CENTER].[dbo].[Item] AS It
								inner JOIN [DATA_CENTER].[dbo].[department] AS Dep ON It.DepartmentID = Dep.ID AND It.storeid = Dep.storeid
								left JOIN [DATA_CENTER].[dbo].[SupplierList] AS SuppL ON It.storeid = SuppL.storeid 
								AND It.SupplierID = SuppL.SupplierID AND It.ID = SuppL.ItemID
								left JOIN [DATA_CENTER].[dbo].[Supplier] AS Supp ON SuppL.SupplierID = Supp.ID AND SuppL.storeid = Supp.storeid
								left join (select RMSstoNumber,Username,DManager,FManager from CkproUsers.dbo.Storeuser) sty on sty.RMSstoNumber =convert(varchar(10),it.storeid)
								left join [DATA_CENTER].dbo.STORES st on st.STORE_ID =it.storeid
								where st.Franchise='SUB-FRANCHISE' and sty.RMSstoNumber !='58'";
        public string DyRptStoreView = @"
select st.Franchise StoreFranchise,st.Storenumber StoreID,
st.username StoreName,
Inv.Modifieddatetime Modified, 
Inv.Wmslocationid StoreNameInDy, 
Inv.Physicalinvent Qty,t.COSTPRICE Cost, Inv.Itemid ItemLookupCode,
It.Name ItemName, CateN.CODE DpId,CateN.Name dpName
,ca.Primaryvendorid SupplierCode,W.SupplierName
,DManager,FManager,Username
 from AXDB.dbo.Inventsum Inv
left join AXDB.dbo.Inventtable ca on Inv.Itemid = ca.Itemid
left join (Select distinct SupplierName,Code from [192.168.1.156].DATA_CENTER.dbo.supplier) w on w.Code=ca.Primaryvendorid collate SQL_Latin1_General_CP1_CI_AS
left join AXDB.dbo.Ecoresproducttranslation It on ca.Product = It.Product
left join AXDB.dbo.Ecoresproductcategory Re on It.Product = Re.Product
left join AXDB.dbo.Ecorescategory CateN on Re.Category = CateN.Recid
left join [192.168.1.156].CkproUsers.dbo.Storeuser st on st.Inventlocation=Inv.Wmslocationid
left join (SELECT distinct s.COSTPRICE,s.ITEMID a FROM  AXDB.dbo.Salesline s WHERE s.Confirmeddlv = (
        SELECT  MAX(Confirmeddlv) FROM  AXDB.dbo.Salesline where itemid=s.ITEMID) )t on t.a=inv.ITEMID
where ca.Dataareaid = 'tmt' ";
        public string RptStockofBranchView = @"select  Inv.INVENTLOCATIONID StoreName,It.Name ItemName,Inv.ItemId ItemLookupcode,Cat.NAME dpName,Cat.CODE dpId, sum(Inv.Physicalinvent) Qty
										,''Username,''DManager,''FManager from [192.168.1.210].AXDB.dbo.INVENTSUM Inv
										inner join [192.168.1.210].AXDB.dbo.Inventtable Ca on Ca.Itemid=Inv.itemid
										inner join [192.168.1.210].AXDB.dbo.Ecoresproducttranslation It on ca.Product = It.Product
										inner join [192.168.1.210].AXDB.dbo.Ecoresproductcategory Re on It.product =Re.Product
										left join [192.168.1.210].AXDB.dbo.Ecorescategory CateN on Re.Category = CateN.Recid
										INNER JOIN  (Select distinct Cat.ReCId,Cat.CODE,Cat.NAME from [192.168.1.210].AXDB.dbo.Ecorescategory Cat)Cat on Re.Category = Cat.Recid

										 where ca.Dataareaid = 'tmt' and  Inv.Wmslocationid is not null or Inv.Wmslocationid=''
										 group by Inv.INVENTLOCATIONID,It.Name,Inv.ItemId,Cat.NAME,Cat.CODE";
         public string RptStoreAll()
        {
              string RmsRptStoreView = @"
								SELECT st.Franchise StoreFranchise,it.StoreID,sty.Username StoreName,
								Dep.code DpId, Dep.Name dpName,It.ItemLookupCode, It.Description ItemName,It.Cost--, It.Price 
								, It.Quantity Qty
								,Supp.Code SupplierCode ,Supp.SupplierName,sty.Username,sty.DManager,sty.FManager
								FROM [DATA_CENTER].[dbo].[Item] AS It
								inner JOIN [DATA_CENTER].[dbo].[department] AS Dep ON It.DepartmentID = Dep.ID AND It.storeid = Dep.storeid
								left JOIN [DATA_CENTER].[dbo].[SupplierList] AS SuppL ON It.storeid = SuppL.storeid 
								AND It.SupplierID = SuppL.SupplierID AND It.ID = SuppL.ItemID
								left JOIN [DATA_CENTER].[dbo].[Supplier] AS Supp ON SuppL.SupplierID = Supp.ID AND SuppL.storeid = Supp.storeid
								left join (select RMSstoNumber,Username,DManager,FManager from CkproUsers.dbo.Storeuser) sty on sty.RMSstoNumber =convert(varchar(10),it.storeid)
								left join [DATA_CENTER].dbo.STORES st on st.STORE_ID =it.storeid
								where st.Franchise='SUB-FRANCHISE' and sty.RMSstoNumber !='58'";
              string linkDyRptStoreView = @"
select st.Franchise StoreFranchise,st.Storenumber StoreID,
st.username StoreName,
Inv.Modifieddatetime Modified, 
Inv.Wmslocationid StoreNameInDy, 
Inv.Physicalinvent Qty,t.COSTPRICE Cost, Inv.Itemid ItemLookupCode,
It.Name ItemName, CateN.CODE DpId,CateN.Name dpName
,ca.Primaryvendorid SupplierCode,W.SupplierName
,DManager,FManager,Username
 from [192.168.1.210].AXDB.dbo.Inventsum Inv
left join [192.168.1.210].AXDB.dbo.Inventtable ca on Inv.Itemid = ca.Itemid
left join [192.168.1.156].DATA_CENTER.dbo.supplier w on w.Code=ca.Primaryvendorid collate SQL_Latin1_General_CP1_CI_AS
left join [192.168.1.210].AXDB.dbo.Ecoresproducttranslation It on ca.Product = It.Product
left join [192.168.1.210].AXDB.dbo.Ecoresproductcategory Re on It.Product = Re.Product
left join [192.168.1.210].AXDB.dbo.Ecorescategory CateN on Re.Category = CateN.Recid
left join [192.168.1.156].CkproUsers.dbo.Storeuser st on st.Inventlocation=Inv.Wmslocationid
left join (SELECT distinct s.COSTPRICE,s.ITEMID a FROM  [192.168.1.210].AXDB.dbo.Salesline s WHERE s.Confirmeddlv = (
        SELECT  MAX(Confirmeddlv) FROM  [192.168.1.210].AXDB.dbo.Salesline where itemid=s.ITEMID) )t on t.a=inv.ITEMID
where ca.Dataareaid = 'tmt'";
              string RptStoreAll = @"
                SELECT StoreFranchise Collate SQL_Latin1_General_CP1_CI_AS StoreFranchise,Convert (varchar(10),StoreID)StoreId,StoreID StoreIdR,''StoreIdD,StoreName Collate SQL_Latin1_General_CP1_CI_AS StoreName,--st.LOCATION StoreName,
                DpId Collate SQL_Latin1_General_CP1_CI_AS DpId, dpName Collate SQL_Latin1_General_CP1_CI_AS dpName,ItemLookupCode Collate SQL_Latin1_General_CP1_CI_AS ItemLookupCode, ItemName Collate SQL_Latin1_General_CP1_CI_AS ItemName,Cost--, Price 
                , Qty
                ,SupplierCode Collate SQL_Latin1_General_CP1_CI_AS SupplierCode ,SupplierName Collate SQL_Latin1_General_CP1_CI_AS SupplierName
                ,Username Collate SQL_Latin1_General_CP1_CI_AS Username,DManager Collate SQL_Latin1_General_CP1_CI_AS DManager,FManager Collate SQL_Latin1_General_CP1_CI_AS FManager 
                from " + $"({RmsRptStoreView})RptStore"+@"
                union all
                SELECT StoreFranchise,StoreID,'',StoreID,StoreName,
                DpId, dpName,ItemLookupCode, ItemName,Cost
                , Qty
                ,SupplierCode ,SupplierName
                ,Username,DManager,FManager
                from "+$"({linkDyRptStoreView})RptAXDBStore";
            return RptStoreAll;
        }
    }
}
//change suppliercode, transactionnumber,storeid,dpid to string in rptsales and 2 and all 