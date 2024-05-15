//using CK.Model;
//using CK.Models;
//using Microsoft.AspNetCore.Authentication;
//using Microsoft.AspNetCore.Authentication.Cookies;
//using Microsoft.AspNetCore.Authorization;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.EntityFrameworkCore;
//using OfficeOpenXml;
//using System.Diagnostics;
//using System.Globalization;
//using Newtonsoft.Json;
//using OfficeOpenXml.Style;
//using Polly;
//using System;
//using ClosedXML.Excel;
//using System.Drawing;
//using Microsoft.CodeAnalysis.Elfie.Model.Structures;
//using System.Linq;
//using Microsoft.CodeAnalysis.CSharp.Syntax;
//using System.Text;
//using Microsoft.Data.SqlClient;
//using DocumentFormat.OpenXml.Features;
//using static System.Runtime.InteropServices.JavaScript.JSType;
//using Microsoft.IdentityModel.Tokens;
//using System.Data;
//using DocumentFormat.OpenXml.InkML;
//using System.Security.Cryptography;
//using System.Windows.Forms;
//namespace CK.Controllers
//{
//    [Authorize]
//    public class HomeController : Controller
//    {
//        private readonly ILogger<HomeController> _logger;
//        public HomeController(ILogger<HomeController> logger)
//        {
//            _logger = logger;
//        }
//        [HttpGet]
//        public IActionResult ATransferSol1(string transferId)
//        {
//            DataCenterContext db = new DataCenterContext();
//            IQueryable<CK.Models.R> reportData1 = Enumerable.Empty<CK.Models.R>().AsQueryable(); // Initialize with an empty queryable
//            var username = HttpContext.Session.GetString("Username");
//            var Role = HttpContext.Session.GetString("Role");
//            ViewBag.Username = username;
//            ViewBag.Role = Role;
//            if (!string.IsNullOrEmpty(transferId))
//            {
//                reportData1 = db.Rs
//                    .Where(d => d.Transaction == transferId)
//                    .GroupBy(d => new { d.FromTable, d.ToTable, d.Item, d.Barcode, d.Day, d.Transaction, d.Status })
//                    .Select(g => new CK.Models.R
//                    {
//                        FromTable = g.Key.FromTable,
//                        Quantity = g.Sum(d => d.Quantity),
//                        Item = g.Key.Item,
//                        Barcode = g.Select(d => d.Barcode).FirstOrDefault(),
//                        Day = g.Key.Day,
//                        Transaction = g.Key.Transaction,
//                        ToTable = g.Key.ToTable,
//                        Status = g.Key.Status
//                    });

//            }

//            // Convert IQueryable to List before passing it to the view
//            var reportDataList = reportData1.ToList();

//            return View(reportDataList);
//        }
//        public IActionResult ATransferSol(string transferId)
//        {
//            DataCenterContext db = new DataCenterContext();
//            IQueryable<CK.Models.R> reportData1 = Enumerable.Empty<CK.Models.R>().AsQueryable(); // Initialize with an empty queryable
//            var username = HttpContext.Session.GetString("Username");
//            var Role = HttpContext.Session.GetString("Role");
//            ViewBag.Username = username;
//            ViewBag.Role = Role;
//            if (!string.IsNullOrEmpty(transferId))
//            {
//                reportData1 = db.Rs
//                    .Where(d => d.Transaction == transferId)
//                    .GroupBy(d => new { d.FromTable, d.ToTable, d.Item, d.Barcode, d.Day, d.Transaction, d.Status })
//                    .Select(g => new CK.Models.R
//                    {
//                        FromTable = g.Key.FromTable,
//                        Quantity = g.Sum(d => d.Quantity),
//                        Item = g.Key.Item,
//                        Barcode = g.Select(d => d.Barcode).FirstOrDefault(),
//                        Day = g.Key.Day,
//                        Transaction = g.Key.Transaction,
//                        ToTable = g.Key.ToTable,
//                        Status = g.Key.Status
//                    });

//            }

//            // Convert IQueryable to List before passing it to the view
//            var reportDataList = reportData1.ToList();

//            return View(reportDataList);
//        }
//        public IActionResult Home()
//        {
//            var username = HttpContext.Session.GetString("Username");
//            var Role = HttpContext.Session.GetString("Role");
//            ViewBag.Username = username;
//            ViewBag.Role = Role;
//            return View();
//        }
//        bool exported = false;
//        [HttpGet]
//        public IActionResult Index()
//        {
//            DataCenterContext db = new DataCenterContext();
//            CkproUsersContext db2 = new CkproUsersContext();
//            CkhelperdbContext db3 = new CkhelperdbContext();
//            AxdbContext Axdb = new AxdbContext();
//            DataCenterPrevYrsContext db4 = new DataCenterPrevYrsContext();
//            db.Database.SetCommandTimeout(7200);// Set the timeout in seconds
//            //Store List Text=StoreName , Value = StoreId
//            var username = HttpContext.Session.GetString("Username");
//            var Role = HttpContext.Session.GetString("Role");
//            ViewBag.Username = username;
//            ViewBag.Role = Role;

//            if (username is null)
//            {
//                return RedirectToAction("Store", "Home");
//            }
//            bool isDmanager = db2.RptUsers.Any(s => s.Dmanager == username);
//            bool isFmanager = db2.RptUsers.Any(s => s.Fmanager == username);
//            bool isUsername = db2.RptUsers.Any(s => s.Username == username && (s.Storenumber != null || s.Storenumber != " "));

//            IQueryable<Storeuser> query;
//            if (isDmanager || isUsername || isFmanager)
//            {
//                // If the username matches either the Dmanager or the Username, filter the stores accordingly
//                query = db2.Storeusers
//                    .Where(s => s.Dmanager == username || s.Username == username || s.Fmanager == username);
//            }
//            else
//            {
//                // If neither condition is met, display all stores
//                query = db2.Storeusers;
//            }
//            ViewBag.VBStore = query
//    .GroupBy(m => m.Name)
//    .Select(group => new { Store = group.First().Storenumber + ":" + group.First().RmsstoNumber, Name = group.Key })
//    .OrderBy(m => m.Name)
//    .ToList();
//            ViewBag.VBDepartment = Axdb.Ecorescategories
//                                                 .GroupBy(m => m.Name)
//                                                 .Select(group => new { Code = group.First().Code, Name = group.Key })
//                                                 .OrderBy(m => m.Name)
//                                                 .ToList();

//            //Supplier List Text=SupplierName , Value = Code 
//            ViewBag.VBSupplier = db.Suppliers
//                                             .GroupBy(m => m.SupplierName)
//                                                 .Select(group => new { Code = group.First().Code, SupplierName = group.Key })
//                                                 .OrderBy(m => m.SupplierName)
//                                                 .ToList();

//            ViewBag.VBItemBarcode = db.Items.Select(m => new { m.Id, m.ItemLookupCode }).Distinct();
//            ViewBag.VBStoreFranchise = db.Stores
//                 .Where(m => m.Franchise != null)
//                 .Select(m => m.Franchise)
//                 .Distinct()
//                 .ToList();
//            return View();
//        }
//        public IActionResult HomeStore()
//        {
//            var username = HttpContext.Session.GetString("Username");
//            ViewBag.Username = username;
//            return View();
//        }
//        [HttpGet]
//        public IActionResult Store()
//        {
//            DataCenterContext db = new DataCenterContext();
//            CkproUsersContext db2 = new CkproUsersContext();
//            CkhelperdbContext db3 = new CkhelperdbContext();
//            DataCenterPrevYrsContext db4 = new DataCenterPrevYrsContext();
//            db.Database.SetCommandTimeout(7200);// Set the timeout in seconds
//            IQueryable<RptSale> RptSales = db.RptSales;
//            IQueryable<RptSalesAxt> RptSalesAxts = db.RptSalesAxts;
//            IQueryable<RptSales2> RptSales2s = db4.RptSales2s;
//            IQueryable<RptSalesAll> RptSalesAlls = db.RptSalesAlls;
//            //Store List Text=StoreName , Value = StoreId
//            var username = HttpContext.Session.GetString("Username");
//            ViewBag.Username = username;
//            bool isDmanager = db2.RptUsers.Any(s => s.Dmanager == username);
//            bool isUsername = db2.RptUsers.Any(s => s.Username == username && (s.Storenumber != null || s.Storenumber != " "));

//            IQueryable<Storeuser> query;
//            if (isDmanager || isUsername)
//            {
//                // If the username matches either the Dmanager or the Username, filter the stores accordingly
//                query = db2.Storeusers
//                    .Where(s => s.Dmanager == username || s.Username == username);
//            }
//            else
//            {
//                // If neither condition is met, display all stores
//                query = db2.Storeusers;
//            }
//            ViewBag.VBStore = query
//    .GroupBy(m => m.Name)
//    .Select(group => new { Store = group.First().Storenumber + ":" + group.First().RmsstoNumber, Name = group.Key })
//    .OrderBy(m => m.Name)
//    .ToList();
//            ViewBag.VBDepartment = db.Departments
//                                                 .GroupBy(m => m.Name)
//                                                 .Select(group => new { Code = group.First().Code, Name = group.Key })
//                                                 .OrderBy(m => m.Name)
//                                                 .ToList();

//            //Supplier List Text=SupplierName , Value = Code 
//            ViewBag.VBSupplier = db.Suppliers
//                                             .GroupBy(m => m.SupplierName)
//                                                 .Select(group => new { Code = group.First().Code, SupplierName = group.Key })
//                                                 .OrderBy(m => m.SupplierName)
//                                                 .ToList();

//            ViewBag.VBItemBarcode = db.Items.Select(m => new { m.Id, m.ItemLookupCode }).Distinct();
//            ViewBag.VBStoreFranchise = db.Stores
//                 .Where(m => m.Franchise != null)
//                 .Select(m => m.Franchise)
//                 .Distinct()
//                 .ToList();
//            return View();
//        }
//        [HttpPost]
//        public IActionResult Store(SalesParameters Parobj)
//        {
//            DataCenterContext db = new DataCenterContext();
//            CkproUsersContext db2 = new CkproUsersContext();
//            CkhelperdbContext db3 = new CkhelperdbContext();
//            DataCenterPrevYrsContext db4 = new DataCenterPrevYrsContext();
//            db.Database.SetCommandTimeout(7200);// Set the timeout in seconds
//            db3.Database.SetCommandTimeout(7200);// Set the timeout in seconds
//            db4.Database.SetCommandTimeout(7200);// Set the timeout in seconds
//            var username = HttpContext.Session.GetString("Username");
//            ViewBag.Username = username;

//            ViewBag.VBStore = db2.Storeusers
//                    .Where(s => s.Dmanager == username || s.Username == username)
//                .GroupBy(m => m.Name)
//                .Select(group => new { Store = group.First().Storenumber + ":" + group.First().RmsstoNumber, Name = group.Key })//group.First().StoreIdD + ":" +
//                .OrderBy(m => m.Name)
//                .ToList();
//            ViewBag.VBDepartment = db.Departments
//                                                 .GroupBy(m => m.Name)
//                                                 .Select(group => new { Code = group.First().Code, Name = group.Key })
//                                                 .OrderBy(m => m.Name)
//                                                 .ToList();

//            //Supplier List Text=SupplierName , Value = Code 
//            ViewBag.VBSupplier = db.Suppliers
//                                             .GroupBy(m => m.SupplierName)
//                                                 .Select(group => new { Code = group.First().Code, SupplierName = group.Key })
//                                                 .OrderBy(m => m.SupplierName)
//                                                 .ToList();

//            ViewBag.VBItemBarcode = db.Items.Select(m => new { m.Id, m.ItemLookupCode }).Distinct();
//            ViewBag.VBStoreFranchise = db.Stores
//                 .Where(m => m.Franchise != null)
//                 .Select(m => m.Franchise)
//                 .Distinct()
//                 .ToList();
//            // Dynamic GroupBy based on selected values
//            IQueryable<dynamic> reportData1;
//            string[] storeVal = Parobj.Store.Split(':');
//            string connectionString1 = string.Format("Server=192.168.1.210;User ID=sa;Password=P@ssw0rd;Database=AXDB;Connect Timeout=7200;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;");
//            string connectionString = string.Format("Server=192.168.1.156;User ID=sa;Password=P@ssw0rd;Database=DATA_CENTER;Connect Timeout=7200;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;");
//            string connectionString2 = string.Format("Server=192.168.1.156;User ID=sa;Password=P@ssw0rd;Database=DATA_CENTER_Prev_Yrs;Connect Timeout=7200;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;");
//            string serverIp = GetServerIpFromDatabase();

//            // Step 2: Format the connection string dynamically
//            //string connectionString = FormatConnectionString(serverIp);
//            // Call the ExportToExcel method
//            Parobj.exportAfterClick = true;
//            Parobj.VQty = true;
//            //Parobj.VTransactionCount = true;
//            Parobj.VTotalSales = true;
//            Parobj.VDepartment = true;
//            Parobj.VStoreName = true;
//            Parobj.RMS = true;
//            Parobj.TMT = true;

//            if (Parobj.RMS || Parobj.TMT)
//            {
//                return ExportToExcel(connectionString, Parobj);
//            }
//            else if (Parobj.DBbefore)
//            {
//                return ExportToExcel(connectionString2, Parobj);
//            }
//            // if Not RMS or TMT
//            else
//            {
//                return View();
//            }
//            ViewBag.Data = reportData1;
//            //  }
//            //TempData["Al"] = "تم الحفظ بفضل الله";
//            //var reportData1 = ViewBag.Data as IEnumerable<dynamic>;
//            if (Parobj.exportAfterClick == false)
//            {
//                return View();
//            }

//            else
//            {
//                // return View();
//                return ExportReportData(reportData1, Parobj);
//            }

//        }

//        public bool IsBase64String(string base64)
//        {
//            Span<byte> buffer = new Span<byte>(new byte[base64.Length]);
//            return Convert.TryFromBase64String(base64, buffer, out _);
//        }
//        public string decrypt(string cipherText)
//        {
//            if (!IsBase64String(cipherText))
//            {
//                // Handle the error, e.g., log it, return a default value, or throw an exception
//                return "Invalid encrypted password format";
//            }

//            string EncryptionKey = "MAKV2SPBNI99212";
//            byte[] cipherBytes = Convert.FromBase64String(cipherText);
//            using (Aes encryptor = Aes.Create())
//            {
//                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
//                encryptor.Key = pdb.GetBytes(32);
//                encryptor.IV = pdb.GetBytes(16);
//                using (MemoryStream ms = new MemoryStream())
//                {
//                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
//                    {
//                        cs.Write(cipherBytes, 0, cipherBytes.Length);
//                        cs.Close();
//                    }
//                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
//                }
//            }
//            return cipherText;
//        }
//        public async Task<IActionResult> indexa()
//        {
//            using (var db2 = new CkproUsersContext())
//            {
//                var storeUsers = await db2.RptUsers2s.ToListAsync();

//                // Decrypt each user's password
//                foreach (var user in storeUsers)
//                {
//                    user.DecryptedPassword = decrypt(user.Password); // Assuming 'decrypt' is your decryption method
//                }

//                return View(storeUsers);
//            }
//        }
//        public async Task<IActionResult> index2()
//        {
//            using (var db2 = new DataCenterContext())
//            {
//                var storeUsers = await db2.RptSales.ToListAsync();

//                return View(storeUsers);
//            }
//        }
//        public IActionResult ExportToExcel(string connectionString, SalesParameters Parobj)
//        {
//            DataCenterContext db = new DataCenterContext();
//            CkproUsersContext db2 = new CkproUsersContext();
//            CkhelperdbContext db3 = new CkhelperdbContext();
//            DataCenterPrevYrsContext db4 = new DataCenterPrevYrsContext();
//            var username = HttpContext.Session.GetString("Username");
//            var Role = HttpContext.Session.GetString("Role");
//            ViewBag.Username = username;
//            ViewBag.Role = Role;
//            HttpContext.Session.SetString("ExportStatus", "started");
//            // Prepare the SQL query with a parameter placeholder
//            // Start building the SELECT clause dynamically
//            List<string> selectColumns = new List<string>();
//            if (Parobj.VPerDay)
//                selectColumns.Add("CAST(transdate as date) as Date");
//            if (Parobj.VPerYear || Parobj.VPerMonYear)
//                selectColumns.Add("ByYear");
//            if (Parobj.VPerMon || Parobj.VPerMonYear)
//                selectColumns.Add("ByMonth");
//            if (Parobj.VStoreId)
//                selectColumns.Add("StoreID");
//            if (Parobj.VStoreName)
//                selectColumns.Add("StoreName as 'Store Name'");
//            if (Parobj.VDpId)
//                selectColumns.Add("DpID as 'Department Id'");
//            if (Parobj.VDepartment)
//                selectColumns.Add("dpname Department");
//            if (Parobj.VItemLookupCode)
//                selectColumns.Add("Itemlookupcode");
//            if (Parobj.VItemName)
//                selectColumns.Add("ItemName");
//            if (Parobj.VSupplierId)
//                selectColumns.Add("SupplierCode");
//            if (Parobj.VSupplierName)
//                selectColumns.Add("SupplierName");
//            if (Parobj.VFranchise)
//                selectColumns.Add("StoreFranchise");
//            if (Parobj.VTransactionNumber)
//                selectColumns.Add("TransactionNumber");
//            if (Parobj.VQty)
//                selectColumns.Add("sum(Qty)TotalQty");
//            if (Parobj.VPrice)
//                selectColumns.Add("Price");
//            if (Parobj.VCost)
//                selectColumns.Add("Cost");
//            if (Parobj.VTotalSales)
//                selectColumns.Add("sum(TotalSales)TotalSales");
//            if (Parobj.VTransactionCount)
//                selectColumns.Add("count(distinct TransactionNumber) TransactionsCount");
//            if (Parobj.VTotalCost)
//                selectColumns.Add("sum(Cost)TotalCost");
//            if (Parobj.VTotalTax)
//                selectColumns.Add("sum(Tax)TotalTax");
//            if (Parobj.VTotalSalesTax)
//                selectColumns.Add("sum(TotalSales)TotalSaleswithTax");
//            if (Parobj.VTotalSalesWithoutTax)
//                selectColumns.Add("sum(TotalSalesWithoutTax)TotalSalesWithoutTax");
//            if (Parobj.VTotalCostQty)
//                selectColumns.Add("sum(TotalCostQty)TotalCostQty");
//            // Construct the SELECT clause from the list of columns
//            string selectClause = string.Join(", ", selectColumns);
//            string fromWhereClause = null;
//            DateTime startDateTime = Convert.ToDateTime(Parobj.startDate, new CultureInfo("en-GB"));
//            DateTime endDateTime = Convert.ToDateTime(Parobj.endDate, new CultureInfo("en-GB"));
//            string[] storeVal = Parobj.Store.Split(':');
//            if ((Parobj.RMS && Parobj.TMT == false) || (Parobj.RMS && storeVal[0] == "RMS"))
//            {
//                if (Parobj.VSupplierId || Parobj.VSupplierName)
//                {
//                    fromWhereClause = "FROM  RptSalesbySupplier WHERE CAST(TransDate AS DATE) BETWEEN @fromDate AND @toDate ";
//                }
//                else
//                {
//                    fromWhereClause = @"FROM
//(select 
//dp.code DpID,dp.ID GroupId, dp.Name DpName,
//T.StoreID StoreID,
//T.ItemID ItemID,Item.Description ItemName,Item.ItemLookupCode,
//T.TransactionTime TransTime,
//Day(T.TransactionTime) ByDay,Month(T.TransactionTime) ByMonth,Year(T.TransactionTime)ByYear, Cast(T.TransactionTime as date) TransDate,T.Quantity Qty,T.Price Price,(T.Quantity * T.Price) TotalSales,
//convert(varchar(20),T.TransactionNumber)TransactionNumber,T.Cost,-((T.Cost)*-(T.Quantity)) TotalCostQty,((T.Quantity*T.Price)-(T.Cost*T.Quantity))Profit,T.SalesTax Tax,
//(T.Quantity*T.Price)TotalSalesTax,((T.Quantity*T.Price)-T.SalesTax)TotalSalesWithoutTax,((T.Quantity*T.Cost)-T.SalesTax)TotalCostWithoutTax
//,Convert (varchar(20),item.SupplierId )SupplierCode1 ";
//                    if (Parobj.VStoreName)
//                    {
//                        fromWhereClause += @" 
//,Stores.ID StoreCode ,Stores.LOCATION StoreName,Stores.FRANCHISE StoreFranchise from TransactionEntry as T )E";
//                    }
//                    else if (Parobj.VStoreName == false)
//                    {
//                        fromWhereClause += " from TransactionEntry as T ";
//                    }
//                    if (Parobj.VStoreName)
//                    {
//                        fromWhereClause += @" 
//left join CkproUsers.dbo.Storeuser s on s.RMSstoNumber=Convert (varchar(20),T.StoreID ))RptSales
//";
//                    }
//                    else
//                    {
//                        fromWhereClause += @" 
//inner join Item on Item.ID=T.ItemID ";
//                        if (Parobj.VStoreName)
//                        {
//                            fromWhereClause += " and T.StoreID=Item.storeid --and item.Quantity!=0";
//                        }
//                        fromWhereClause += " left join department dp on dp.ID=Item.DepartmentID ";
//                    }
//                    if (Parobj.VStoreName)
//                    {
//                        fromWhereClause += " and dp.storeid=Item.storeid left join Stores on (Stores.STORE_ID=Item.storeid or Item.storeid is null) and Stores.STORE_ID=T.StoreID";
//                    }
//                    if (Parobj.startDate != null)
//                    {
//                        fromWhereClause += ")R WHERE CAST(TransDate AS DATE) BETWEEN @fromDate AND @toDate ";
//                    }

//                }
//            }
//            else if (Parobj.RMS == false && Parobj.TMT || (storeVal.Length > 1 && storeVal[1] == "Dy"))
//            {
//                fromWhereClause = "FROM RptSalesData where CAST(TransDate AS DATE) BETWEEN @fromDate AND @toDate ";
//                //fromWhereClause ="from (Select SalesD.TRANSACTIONID TransactionNumber,SalesH.STORE StoreID,SalesD.COSTAMOUNT Cost--,Store.Name StoreName,SalesD.ITEMID ItemLookupCode,-TAXAMOUNT Tax,Day(SalesH.TRANSDATE) ByDay,Month(SalesH.TRANSDATE) ByMonth,Year(SalesH.TRANSDATE)ByYear,SalesH.TRANSDATE TransTime,SalesH.TRANSDATE As TransDate,-SalesD.Qty Qty,-(SalesD.COSTAMOUNT*SalesD.Qty)TotalCostQty,SalesD.Price Price,-(SalesD.NETAMOUNTINCLTAX) TotalSales,-(SalesD.NETAMOUNTINCLTAX) TotalSalesTax, -(SalesD.NETAMOUNT) TotalSalesWithoutTax,Case when SalesH.STORE='143' then 'Sub-Franchise' else 'TMT' end As StoreFranchise ,INV.[Primaryvendorid] SupplierName, INV.[Primaryvendorid] SupplierCode ,It.[NAME] ItemName from  RetailChannelDatabase.ax.RetailTransactiontable SalesH INNER JOIN RetailChannelDatabase.ax.RETAILTRANSACTIONSALESTRANS SalesD ON SalesH.TRANSACTIONID = SalesD.TRANSACTIONID INNER JOIN  RetailChannelDatabase.ax.[Inventtable] as INV on SalesD.ITEMID = INV.ITEMID inner JOIN  RetailChannelDatabase.ax.[Ecoresproducttranslation] as It on INV.PRODUCT = It.PRODUCT   where SalesH.ENTRYSTATUS!=1  and SalesH.TYPE=2 AND INV.[DATAAREAID] = 'tmt' )s";            }
//            }
//            //if (Parobj.RMS && Parobj.TMT && storeVal[0] != "RMS" && storeVal[1] != "Dy" || (Parobj.RMS && Parobj.TMT && storeVal[0] == "0" && storeVal[1] == "0"))
//            else if (Parobj.DBbefore)
//            {
//                // fromWhereClause = "FROM AAArptsales where CAST(TransDate AS DATE) BETWEEN @fromDate AND @toDate ";
//                fromWhereClause = "FROM RptSales2 where CAST(TransDate AS DATE) BETWEEN @fromDate AND @toDate ";
//            }
//            else
//            {
//                if (Parobj.VSupplierId || Parobj.VSupplierName)
//                {
//                    fromWhereClause = "from RptSalesAllwithSuppllier WHERE CAST(TransDate AS DATE) BETWEEN @fromDate AND @toDate ";
//                }
//                else
//                {
//                    fromWhereClause = "from RptSalesAll WHERE CAST(TransDate AS DATE) BETWEEN @fromDate AND @toDate ";
//                }
//            }
//            // string MessageBox = string.Empty;
//            // Add department filter if a department is specified
//            if (!string.IsNullOrEmpty(Parobj.Department) && Parobj.Department != "0")
//            {
//                fromWhereClause += "AND DpId = @Department ";
//            }
//            if (!string.IsNullOrEmpty(Parobj.Supplier) && Parobj.Supplier != "0")
//            {
//                fromWhereClause += "AND SupplierCode = @Supplier ";
//            }
//            bool isFmanager = db2.RptUsers.Any(s => s.Fmanager == username);
//            bool isDmanager = db2.RptUsers.Any(s => s.Dmanager == username);
//            bool isUsername = db2.RptUsers.Any(s => s.Username == username && (s.Storenumber != null || s.Storenumber != " "));
//            IQueryable<Storeuser> query;
//            if (isDmanager || isUsername || isFmanager)
//            {
//                fromWhereClause += "AND (Dmanager='" + username + "' or username ='" + username + "' or Fmanager ='" + username + "') ";

//            }
//            if (Parobj.Store != "0")
//            {
//                if (Parobj.RMS && Parobj.TMT == false || storeVal[0] == "RMS" || Parobj.DBbefore)
//                {
//                    fromWhereClause += "AND StoreId = @Store1 ";
//                }
//                else if (Parobj.RMS == false && Parobj.TMT || (storeVal.Length > 1 && storeVal[1] == "Dy"))
//                {
//                    fromWhereClause += "AND StoreId = @Store ";
//                }
//                else if (Parobj.RMS && Parobj.TMT)
//                {
//                    fromWhereClause += "AND (StoreIdD = @Store OR StoreIdR = @Store1) ";

//                }
//                else
//                {
//                    fromWhereClause += "AND StoreId = @Store1 ";
//                }
//            }
//            if (Parobj.Franchise == "TMT")
//            {
//                fromWhereClause += "AND StoreFranchise = 'TMT'";
//            }
//            if (Parobj.Franchise == "SUB-FRANCHISE")
//            {
//                fromWhereClause += "AND StoreFranchise = 'SUB-FRANCHISE'";
//            }
//            // Add ItemLookupCode filter if ItemLookupCodeTxt is not null or empty
//            if (!string.IsNullOrEmpty(Parobj.ItemLookupCodeTxt))
//            {
//                // Split the string into an array of values
//                string[] itemLookupCodes = Parobj.ItemLookupCodeTxt.Split(',');

//                // Start building the IN clause
//                fromWhereClause += " AND ItemLookupCode IN (";

//                // Add a parameter placeholder for each value
//                for (int i = 0; i < itemLookupCodes.Length; i++)
//                {
//                    // Trim any whitespace from the value
//                    string itemLookupCode = itemLookupCodes[i].Trim();

//                    // Append the parameter placeholder to the IN clause
//                    fromWhereClause += $"@ItemLookupCode{i}";

//                    // Add a comma separator if not the last item
//                    if (i < itemLookupCodes.Length - 1)
//                    {
//                        fromWhereClause += ",";
//                    }
//                }

//                fromWhereClause += ")";
//            }
//            if (!string.IsNullOrEmpty(Parobj.ItemNameTxt))
//            {
//                // Split the string into an array of values
//                string[] ItemNames = Parobj.ItemNameTxt.Split(',');

//                // Start building the IN clause
//                fromWhereClause += " AND ItemLookupCode IN (";

//                // Add a parameter placeholder for each value
//                for (int i = 0; i < ItemNames.Length; i++)
//                {
//                    // Trim any whitespace from the value
//                    string ItemName = ItemNames[i].Trim();

//                    // Append the parameter placeholder to the IN clause
//                    fromWhereClause += $"@ItemName{i}";

//                    // Add a comma separator if not the last item
//                    if (i < ItemNames.Length - 1)
//                    {
//                        fromWhereClause += ",";
//                    }
//                }

//                fromWhereClause += ")";
//            }
//            string sqlQuery = $"SELECT {selectClause} {fromWhereClause}";
//            // Start building the GROUP BY clause dynamically
//            List<string> groupByColumns = new List<string>();
//            if (Parobj.VPerYear || Parobj.VPerMonYear)
//                groupByColumns.Add("ByYear");
//            if (Parobj.VPerMon || Parobj.VPerMonYear)
//                groupByColumns.Add("ByMonth");
//            if (Parobj.VStoreId)
//                groupByColumns.Add("StoreID");
//            if (Parobj.VStoreName)
//                groupByColumns.Add("StoreName");
//            if (Parobj.VDpId)
//                groupByColumns.Add("DpID");
//            if (Parobj.VDepartment)
//                groupByColumns.Add("dpname");
//            if (Parobj.VItemLookupCode)
//                groupByColumns.Add("Itemlookupcode");
//            if (Parobj.VItemName)
//                groupByColumns.Add("ItemName");
//            if (Parobj.VSupplierId)
//                groupByColumns.Add("SupplierCode");
//            if (Parobj.VSupplierName)
//                groupByColumns.Add("SupplierName");
//            if (Parobj.VFranchise)
//                groupByColumns.Add("StoreFranchise");
//            if (Parobj.VTransactionNumber)
//                groupByColumns.Add("TransactionNumber");
//            if (Parobj.VPrice)
//                groupByColumns.Add("Price");
//            if (Parobj.VCost)
//                groupByColumns.Add("Cost");
//            if (Parobj.VPerDay)
//                groupByColumns.Add("CAST(transdate as date)");
//            // Do not include sum(totalsales) in the GROUP BY clause

//            // Construct the GROUP BY clause from the list of columns
//            string groupByClause = groupByColumns.Count > 0 ? "GROUP BY " + string.Join(", ", groupByColumns) : "";

//            sqlQuery += groupByClause;
//            using (SqlConnection connection = new SqlConnection(connectionString))
//            {
//                connection.Open();
//                //string storedProcedureName = "R2"; // Replace with your actual stored procedure name
//                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
//                {
//                    //command.CommandType = CommandType.StoredProcedure;
//                    //sqlQuery = "SELECT sum(TotalSales) TotalSales FROM RptSales WHERE CAST(TransDate AS DATE) BETWEEN @fromDate AND @toDate";
//                    command.CommandTimeout = 7200;
//                    // Add the date parameters to the command if they are not null
//                    if (!string.IsNullOrEmpty(Parobj.startDate))
//                    {
//                        command.Parameters.AddWithValue("@fromDate", startDateTime.Date.ToString("yyyy-MM-dd"));
//                    }
//                    if (!string.IsNullOrEmpty(Parobj.endDate))
//                    {
//                        command.Parameters.AddWithValue("@toDate", endDateTime.Date.ToString("yyyy-MM-dd"));
//                    }
//                    if (Parobj.Store != "0")
//                    {
//                        if (Parobj.RMS && Parobj.TMT == false || storeVal[0] == "RMS")
//                        {
//                            if (storeVal.Length > 1 && int.TryParse(storeVal[1], out int storeId))
//                            {
//                                command.Parameters.AddWithValue("@Store1", storeId);
//                            }
//                        }
//                        else if (Parobj.RMS == false && Parobj.TMT || (storeVal.Length > 1 && storeVal[1] == "Dy"))
//                        {
//                            if (storeVal.Length > 1 && int.TryParse(storeVal[0], out int storeId))
//                            {
//                                command.Parameters.AddWithValue("@Store", storeId);
//                            }
//                        }
//                        else if (Parobj.RMS && Parobj.TMT)// && storeVal[0] != "RMS" && storeVal[1] != "Dy" || (Parobj.RMS && Parobj.TMT && storeVal[0] == "0" && storeVal[1] == "0"))
//                        {
//                            if (storeVal.Length > 1)
//                            {
//                                int storeIdd, storeIdr;
//                                // Attempt to parse storeVal[0] and storeVal[1] as integers
//                                bool isStoreIddParsed = int.TryParse(storeVal[0], out storeIdd);
//                                bool isStoreIdrParsed = int.TryParse(storeVal[1], out storeIdr);

//                                // Check if at least one of the values was successfully parsed
//                                if (isStoreIddParsed || isStoreIdrParsed)
//                                {
//                                    // If storeVal[0] was successfully parsed, add it as a parameter
//                                    if (isStoreIddParsed)
//                                    {
//                                        command.Parameters.AddWithValue("@Store", storeIdd);
//                                    }

//                                    // If storeVal[1] was successfully parsed, add it as a parameter
//                                    if (isStoreIdrParsed)
//                                    {
//                                        command.Parameters.AddWithValue("@Store1", storeIdr);
//                                    }
//                                }
//                            }
//                        }
//                        else
//                        {
//                            return View();
//                        }
//                    }
//                    if (!string.IsNullOrEmpty(Parobj.Department) && Parobj.Department != "0")
//                    {
//                        command.Parameters.AddWithValue("@Department", Parobj.Department);
//                    }
//                    if (!string.IsNullOrEmpty(Parobj.Supplier) && Parobj.Supplier != "0")
//                    {
//                        command.Parameters.AddWithValue("@Supplier", Parobj.Supplier);
//                    }
//                    if (!string.IsNullOrEmpty(Parobj.ItemLookupCodeTxt))
//                    {
//                        string[] itemLookupCodes = Parobj.ItemLookupCodeTxt.Split(',');
//                        for (int i = 0; i < itemLookupCodes.Length; i++)
//                        {
//                            string itemLookupCode = itemLookupCodes[i].Trim();
//                            command.Parameters.AddWithValue($"@ItemLookupCode{i}", itemLookupCode);
//                        }
//                    }
//                    if (!string.IsNullOrEmpty(Parobj.ItemNameTxt))
//                    {
//                        string[] ItemNames = Parobj.ItemNameTxt.Split(',');
//                        for (int i = 0; i < ItemNames.Length; i++)
//                        {
//                            string ItemName = ItemNames[i].Trim();
//                            command.Parameters.AddWithValue($"@ItemName{i}", ItemName);
//                        }
//                    }
//                    // Create a new Excel package
//                    //     try
//                    //    {
//                    using (var package = new ExcelPackage())
//                    {
//                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("AKSalesReport");
//                        int row = 2; // Start from row 2 to leave space for headers
//                        int sheetIndex = 1; // Start with the first sheet
//                        int columnCount = 1;
//                        // Add header row
//                        void AddHeaderRow(ExcelWorksheet ws, int columnCount)
//                        {
//                            int column = 1;
//                            if (Parobj.VPerYear || Parobj.VPerMonYear)
//                                ws.Cells[1, column++].Value = "ByYear";
//                            if (Parobj.VPerMon || Parobj.VPerMonYear)
//                                ws.Cells[1, column++].Value = "ByMonth";
//                            if (Parobj.VStoreId)
//                                ws.Cells[1, column++].Value = "StoreID";
//                            if (Parobj.VStoreName)
//                                ws.Cells[1, column++].Value = "Store Name";
//                            if (Parobj.VDpId)
//                                ws.Cells[1, column++].Value = "Department Id";
//                            if (Parobj.VDepartment)
//                                ws.Cells[1, column++].Value = "Department";
//                            if (Parobj.VItemLookupCode)
//                                ws.Cells[1, column++].Value = "BarCode";
//                            if (Parobj.VItemName)
//                                ws.Cells[1, column++].Value = "ItemName";
//                            if (Parobj.VSupplierId)
//                                ws.Cells[1, column++].Value = "SupplierCode";
//                            if (Parobj.VSupplierName)
//                                ws.Cells[1, column++].Value = "SupplierName";
//                            if (Parobj.VFranchise)
//                                ws.Cells[1, column++].Value = "StoreFranchise";
//                            if (Parobj.VTransactionNumber)
//                                ws.Cells[1, column++].Value = "TransactionNumber";
//                            if (Parobj.VQty)
//                                ws.Cells[1, column++].Value = "TotalQty";
//                            if (Parobj.VPrice)
//                                ws.Cells[1, column++].Value = "Price";
//                            if (Parobj.VCost)
//                                ws.Cells[1, column++].Value = "Cost";
//                            if (Parobj.VTotalSales)
//                                ws.Cells[1, column++].Value = "TotalSales";
//                            if (Parobj.VTransactionCount)
//                                ws.Cells[1, column++].Value = "TransactionsCount";
//                            if (Parobj.VTotalCost)
//                                ws.Cells[1, column++].Value = "TotalCost";
//                            if (Parobj.VTotalTax)
//                                ws.Cells[1, column++].Value = "TotalTax";
//                            if (Parobj.VTotalSalesTax)
//                                ws.Cells[1, column++].Value = "TotalSaleswithTax";
//                            if (Parobj.VTotalSalesWithoutTax)
//                                ws.Cells[1, column++].Value = "TotalSalesWithoutTax";
//                            if (Parobj.VTotalCostQty)
//                                ws.Cells[1, column++].Value = "TotalCostQty";
//                            if (Parobj.VPerDay)
//                                ws.Cells[1, column++].Value = "Date";
//                            using (var headerRange = ws.Cells[1, 1, 1, column - 1])
//                            {
//                                headerRange.Style.Font.Bold = true;
//                                headerRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
//                                headerRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
//                                headerRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
//                                headerRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
//                                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
//                                headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
//                                headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SkyBlue);
//                                ws.Cells[1, 1, 1, column - 1].AutoFilter = true;
//                            }
//                        }
//                        AddHeaderRow(worksheet, columnCount);
//                        //row = 2;
//                        using (SqlDataReader reader = command.ExecuteReader())
//                        {
//                            while (reader.Read())
//                            {
//                                columnCount = 1; // Reset column count for each row
//                                if (Parobj.VPerYear || Parobj.VPerMonYear)
//                                    worksheet.Cells[row, columnCount++].Value = reader["ByYear"];
//                                if (Parobj.VPerMon || Parobj.VPerMonYear)
//                                    worksheet.Cells[row, columnCount++].Value = reader["ByMonth"];
//                                if (Parobj.VStoreId)
//                                    worksheet.Cells[row, columnCount++].Value = reader["StoreID"];
//                                if (Parobj.VStoreName)
//                                    worksheet.Cells[row, columnCount++].Value = reader["Store Name"];
//                                if (Parobj.VDpId)
//                                    worksheet.Cells[row, columnCount++].Value = reader["Department Id"];
//                                if (Parobj.VDepartment)
//                                    worksheet.Cells[row, columnCount++].Value = reader["Department"];
//                                if (Parobj.VItemLookupCode)
//                                    worksheet.Cells[row, columnCount++].Value = reader["Itemlookupcode"];
//                                if (Parobj.VItemName)
//                                    worksheet.Cells[row, columnCount++].Value = reader["ItemName"];
//                                if (Parobj.VSupplierId)
//                                    worksheet.Cells[row, columnCount++].Value = reader["SupplierCode"];
//                                if (Parobj.VSupplierName)
//                                    worksheet.Cells[row, columnCount++].Value = reader["SupplierName"];
//                                if (Parobj.VFranchise)
//                                    worksheet.Cells[row, columnCount++].Value = reader["StoreFranchise"];
//                                if (Parobj.VTransactionNumber)
//                                    worksheet.Cells[row, columnCount++].Value = reader["TransactionNumber"];
//                                if (Parobj.VQty)
//                                    worksheet.Cells[row, columnCount++].Value = reader["TotalQty"];
//                                if (Parobj.VPrice)
//                                    worksheet.Cells[row, columnCount++].Value = reader["Price"];
//                                if (Parobj.VCost)
//                                    worksheet.Cells[row, columnCount++].Value = reader["Cost"];
//                                worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                                if (Parobj.VTotalSales)
//                                    worksheet.Cells[row, columnCount++].Value = reader["TotalSales"];
//                                if (Parobj.VTransactionCount)
//                                    worksheet.Cells[row, columnCount++].Value = reader["TransactionsCount"];
//                                if (Parobj.VTotalCost)
//                                    worksheet.Cells[row, columnCount++].Value = reader["TotalCost"];
//                                if (Parobj.VTotalTax)
//                                    worksheet.Cells[row, columnCount++].Value = reader["TotalTax"];
//                                if (Parobj.VTotalSalesTax)
//                                    worksheet.Cells[row, columnCount++].Value = reader["TotalSaleswithTax"];
//                                if (Parobj.VTotalSalesWithoutTax)
//                                    worksheet.Cells[row, columnCount++].Value = reader["TotalSalesWithoutTax"];
//                                if (Parobj.VTotalCostQty)
//                                    worksheet.Cells[row, columnCount++].Value = reader["TotalCostQty"];
//                                worksheet.Cells[row, columnCount].Style.Numberformat.Format = "yyyy-MM-dd";
//                                if (Parobj.VPerDay)
//                                    worksheet.Cells[row, columnCount++].Value = reader["Date"];
//                                if (columnCount <= 1)
//                                {
//                                    Console.WriteLine("Error: columnCount is 0. No data to process.");
//                                    // Optionally, throw an exception to halt execution
//                                    // throw new InvalidOperationException("columnCount is 0. No data to process.");
//                                }
//                                else
//                                {
//                                    using (var rowRange = worksheet.Cells[row, 1, row, columnCount - 1])
//                                    {
//                                        rowRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

//                                        if (row % 2 == 0) // Even row
//                                        {
//                                            rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
//                                            rowRange.Style.Fill.BackgroundColor.SetColor(Color.LightBlue); // Light gray for even rows
//                                        }
//                                        rowRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
//                                        rowRange.Style.Border.Top.Color.SetColor(Color.LightBlue); // Set border color to black
//                                        rowRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
//                                        rowRange.Style.Border.Bottom.Color.SetColor(Color.LightBlue); // Set border color to black
//                                        rowRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
//                                        rowRange.Style.Border.Left.Color.SetColor(Color.LightBlue); // Set border color to black
//                                        rowRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
//                                        rowRange.Style.Border.Right.Color.SetColor(Color.LightBlue); // Set border color to black
//                                    }
//                                    row++;
//                                }

//                                if (row == 1000001)
//                                {
//                                    // Create a new worksheet and reset the row count
//                                    worksheet = package.Workbook.Worksheets.Add($"AKSalesReport{sheetIndex++}");
//                                    // Re-add the header row to the new worksheet
//                                    row = 2; // Reset row count for the new worksheet
//                                    columnCount = 1; // Reset column count
//                                                     // Re-add the header row to the new worksheet\
//                                    AddHeaderRow(worksheet, columnCount);
//                                    if (Parobj.VPerYear || Parobj.VPerMonYear)
//                                        worksheet.Cells[row, columnCount++].Value = reader["ByYear"];
//                                    if (Parobj.VPerMon || Parobj.VPerMonYear)
//                                        worksheet.Cells[row, columnCount++].Value = reader["ByMonth"];
//                                    if (Parobj.VStoreId)
//                                        worksheet.Cells[row, columnCount++].Value = reader["StoreID"];
//                                    if (Parobj.VStoreName)
//                                        worksheet.Cells[row, columnCount++].Value = reader["Store Name"];
//                                    if (Parobj.VDpId)
//                                        worksheet.Cells[row, columnCount++].Value = reader["Department Id"];
//                                    if (Parobj.VDepartment)
//                                        worksheet.Cells[row, columnCount++].Value = reader["Department"];
//                                    if (Parobj.VItemLookupCode)
//                                        worksheet.Cells[row, columnCount++].Value = reader["Itemlookupcode"];
//                                    if (Parobj.VItemName)
//                                        worksheet.Cells[row, columnCount++].Value = reader["ItemName"];
//                                    if (Parobj.VSupplierId)
//                                        worksheet.Cells[row, columnCount++].Value = reader["SupplierCode"];
//                                    if (Parobj.VSupplierName)
//                                        worksheet.Cells[row, columnCount++].Value = reader["SupplierName"];
//                                    if (Parobj.VFranchise)
//                                        worksheet.Cells[row, columnCount++].Value = reader["StoreFranchise"];
//                                    if (Parobj.VTransactionNumber)
//                                        worksheet.Cells[row, columnCount++].Value = reader["TransactionNumber"];
//                                    if (Parobj.VQty)
//                                        worksheet.Cells[row, columnCount++].Value = reader["TotalQty"];
//                                    if (Parobj.VPrice)
//                                        worksheet.Cells[row, columnCount++].Value = reader["Price"];
//                                    if (Parobj.VCost)
//                                        worksheet.Cells[row, columnCount++].Value = reader["Cost"];
//                                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                                    if (Parobj.VTotalSales)
//                                        worksheet.Cells[row, columnCount++].Value = reader["TotalSales"];
//                                    if (Parobj.VTransactionCount)
//                                        worksheet.Cells[row, columnCount++].Value = reader["TransactionsCount"];
//                                    if (Parobj.VTotalCost)
//                                        worksheet.Cells[row, columnCount++].Value = reader["TotalCost"];
//                                    if (Parobj.VTotalTax)
//                                        worksheet.Cells[row, columnCount++].Value = reader["TotalTax"];
//                                    if (Parobj.VTotalSalesTax)
//                                        worksheet.Cells[row, columnCount++].Value = reader["TotalSaleswithTax"];
//                                    if (Parobj.VTotalSalesWithoutTax)
//                                        worksheet.Cells[row, columnCount++].Value = reader["TotalSalesWithoutTax"];
//                                    if (Parobj.VTotalCostQty)
//                                        worksheet.Cells[row, columnCount++].Value = reader["TotalCostQty"];
//                                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "yyyy-MM-dd";
//                                    if (Parobj.VPerDay)
//                                        worksheet.Cells[row, columnCount++].Value = reader["Date"];
//                                    if (columnCount <= 1)
//                                    {
//                                        Console.WriteLine("Error: columnCount is 0. No data to process.");
//                                        // Optionally, throw an exception to halt execution
//                                        // throw new InvalidOperationException("columnCount is 0. No data to process.");
//                                    }
//                                    else
//                                    {
//                                        // Apply styling to the row
//                                        using (var rowRange = worksheet.Cells[row, 1, row, columnCount - 1])
//                                        {
//                                            rowRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

//                                            if (row % 2 == 0) // Even row
//                                            {
//                                                rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
//                                                rowRange.Style.Fill.BackgroundColor.SetColor(Color.LightBlue); // Light gray for even rows
//                                            }
//                                            rowRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
//                                            rowRange.Style.Border.Top.Color.SetColor(Color.LightBlue); // Set border color to black
//                                            rowRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
//                                            rowRange.Style.Border.Bottom.Color.SetColor(Color.LightBlue); // Set border color to black
//                                            rowRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
//                                            rowRange.Style.Border.Left.Color.SetColor(Color.LightBlue); // Set border color to black
//                                            rowRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
//                                            rowRange.Style.Border.Right.Color.SetColor(Color.LightBlue); // Set border color to black
//                                        }
//                                    }
//                                }
//                            }
//                        }
//                        worksheet.Cells.AutoFitColumns();
//                        // Save the package to a MemoryStream
//                        var stream = new MemoryStream();
//                        package.SaveAs(stream);

//                        // Reset the stream position to the beginning
//                        stream.Position = 0;
//                        Console.WriteLine(sqlQuery); // Print the final query string

//                        // Before executing the command
//                        foreach (SqlParameter param in command.Parameters)
//                        {
//                            Console.WriteLine($"{param.ParameterName}: {param.Value}"); // Print each parameter name and value
//                        }
//                        // Return the file as a FileResult
//                        Console.WriteLine(sqlQuery); // Print the final query string
//                        foreach (SqlParameter param in command.Parameters)
//                        {
//                            Console.WriteLine($"{param.ParameterName}: {param.Value}"); // Print each parameter name and value
//                        }
//                        HttpContext.Session.SetString("ExportStatus", "complete");
//                        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "AKHelperSales.xlsx");
//                    }
//                    // }
//                    //    catch {
//                    //       HttpContext.Session.SetString("ExportStatus", "unKnown1");
//                    //        return View();
//                    //  }
//                }
//            }
//        }
//        static void Main(string[] args)
//        {
//            // Step 1: Retrieve the server IP from the database
//            string serverIp = GetServerIpFromDatabase();

//            // Step 2: Format the connection string dynamically
//            string connectionString = FormatConnectionString(serverIp);

//            // Use the connection string to connect to the database
//            // For demonstration, let's just print the connection string
//            Console.WriteLine(connectionString);
//        }
//        static string GetServerIpFromDatabase()
//        {
//            string serverIp = string.Empty;
//            string connectionString = "Server=192.168.1.156;User ID=sa;Password=P@ssw0rd;Database=CkproUsers;Connect Timeout=7200;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;";

//            using (SqlConnection connection = new SqlConnection(connectionString))
//            {
//                connection.Open();
//                string query = "SELECT TOP 1 server FROM Storeuser where Server='192.168.104.222/New'"; // Assuming you want the first server IP found
//                using (SqlCommand command = new SqlCommand(query, connection))
//                {
//                    object result = command.ExecuteScalar();
//                    if (result != null)
//                    {
//                        serverIp = result.ToString();
//                    }
//                }
//            }

//            return serverIp;
//        }

//        static string FormatConnectionString(string serverIp)
//        {
//            // Assuming the rest of the connection string remains the same except for the server IP
//            string connectionStringFormat = "Server={0};User ID=sa;Password=P@ssw0rd;Database=RetailChannelDatabase;Connect Timeout=7200;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;";
//            string connectionString = string.Format(connectionStringFormat, serverIp);
//            return connectionString;
//        }
//        [HttpPost]
//        public IActionResult Index(SalesParameters Parobj)
//        {
//            DataCenterContext db = new DataCenterContext();
//            CkproUsersContext db2 = new CkproUsersContext();
//            CkhelperdbContext db3 = new CkhelperdbContext();
//            DataCenterPrevYrsContext db4 = new DataCenterPrevYrsContext();
//            db.Database.SetCommandTimeout(7200);// Set the timeout in seconds
//            db3.Database.SetCommandTimeout(7200);// Set the timeout in seconds
//            db4.Database.SetCommandTimeout(7200);// Set the timeout in seconds
//            var username = HttpContext.Session.GetString("Username");
//            var Role = HttpContext.Session.GetString("Role");
//            ViewBag.Username = username;
//            ViewBag.Role = Role;

//            ViewBag.VBStore = db2.Storeusers
//                    .Where(s => s.Dmanager == username || s.Username == username)
//                .GroupBy(m => m.Name)
//                .Select(group => new { Store = group.First().Storenumber + ":" + group.First().RmsstoNumber, Name = group.Key })//group.First().StoreIdD + ":" +
//                .OrderBy(m => m.Name)
//                .ToList();
//            ViewBag.VBDepartment = db.Departments
//                                                 .GroupBy(m => m.Name)
//                                                 .Select(group => new { Code = group.First().Code, Name = group.Key })
//                                                 .OrderBy(m => m.Name)
//                                                 .ToList();

//            //Supplier List Text=SupplierName , Value = Code 
//            ViewBag.VBSupplier = db.Suppliers
//                                             .GroupBy(m => m.SupplierName)
//                                                 .Select(group => new { Code = group.First().Code, SupplierName = group.Key })
//                                                 .OrderBy(m => m.SupplierName)
//                                                 .ToList();

//            ViewBag.VBItemBarcode = db.Items.Select(m => new { m.Id, m.ItemLookupCode }).Distinct();
//            ViewBag.VBStoreFranchise = db.Stores
//                 .Where(m => m.Franchise != null)
//                 .Select(m => m.Franchise)
//                 .Distinct()
//                 .ToList();
//            // Dynamic GroupBy based on selected values
//            IQueryable<dynamic> reportData1;
//            string[] storeVal = Parobj.Store.Split(':');
//            string connectionString1 = string.Format("Server=192.168.1.210;User ID=sa;Password=P@ssw0rd;Database=AXDB;Connect Timeout=7200;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;");
//            string connectionString = string.Format("Server=192.168.1.156;User ID=sa;Password=P@ssw0rd;Database=DATA_CENTER;Connect Timeout=7200;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;");
//            string connectionString2 = string.Format("Server=192.168.1.156;User ID=sa;Password=P@ssw0rd;Database=DATA_CENTER_Prev_Yrs;Connect Timeout=7200;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False;");
//            string serverIp = GetServerIpFromDatabase();

//            // Step 2: Format the connection string dynamically
//            //string connectionString = FormatConnectionString(serverIp);
//            // Call the ExportToExcel method

//            if (Parobj.RMS || Parobj.TMT)
//            {
//                return ExportToExcel(connectionString, Parobj);
//            }
//            else if (Parobj.DBbefore)
//            {
//                return ExportToExcel(connectionString2, Parobj);
//            }
//            // if Not RMS or TMT
//            else
//            {
//                return View();
//            }
//            ViewBag.Data = reportData1;
//            //  }
//            //TempData["Al"] = "تم الحفظ بفضل الله";
//            //var reportData1 = ViewBag.Data as IEnumerable<dynamic>;
//            Parobj.exportAfterClick = true;
//            if (Parobj.exportAfterClick == false)
//            {
//                return View();
//            }

//            else
//            {
//                // return View();
//                return ExportReportData(reportData1, Parobj);
//            }
//            //TempData["Al"] = "تم الحفظ بفضل الله";


//            //Parobj.exportAfterClick = true;
//            //if (Parobj.exportAfterClick == false)
//            //{
//            //    return View();
//            //}
//            //else if (Parobj.TMT)
//            //{
//            //    return ExportToExcelAx(connectionString, Parobj);
//            //}

//        }
//        public IActionResult CheckExportStatus()
//        {
//            // Read the session variable
//            var exportStatus = HttpContext.Session.GetString("ExportStatus");
//            if (exportStatus == "complete")
//            {
//                // If the status is "complete", reset it to an empty string or any other default value
//                HttpContext.Session.SetString("ExportStatus", "");
//                return Content("complete");
//            }
//            return Content(exportStatus ?? "unknown");
//        }
//        private IActionResult ExportReportData(IEnumerable<dynamic> reportData1, SalesParameters Parobj)
//        {
//            HttpContext.Session.SetString("ExportStatus", "started");
//            using (var package = new ExcelPackage())
//            {
//                var worksheet = package.Workbook.Worksheets.Add("AKSalesReport");
//                // Add header row
//                int columnCount = 1; // Start with the first column (A)

//                if (Parobj.VPerYear || Parobj.VPerMonYear)
//                    worksheet.Cells[1, columnCount++].Value = "Date Per Year";
//                if (Parobj.VPerMon || Parobj.VPerMonYear)
//                    worksheet.Cells[1, columnCount++].Value = "Date Per Month";
//                if (Parobj.VPerDay)
//                    worksheet.Cells[1, columnCount++].Value = "Date Per Day";
//                if (Parobj.VStoreId)
//                    worksheet.Cells[1, columnCount++].Value = "Store Id";
//                if (Parobj.VStoreName)
//                    worksheet.Cells[1, columnCount++].Value = "Store Name";
//                if (Parobj.VDpId)
//                    worksheet.Cells[1, columnCount++].Value = "Department Id";
//                if (Parobj.VDepartment)
//                    worksheet.Cells[1, columnCount++].Value = "Department Name";
//                if (Parobj.VItemLookupCode)
//                    worksheet.Cells[1, columnCount++].Value = "Item Lookup Code";
//                if (Parobj.VItemName)
//                    worksheet.Cells[1, columnCount++].Value = "Item Name";
//                if (Parobj.VSupplierId)
//                    worksheet.Cells[1, columnCount++].Value = "Supplier Code";
//                if (Parobj.VSupplierName)
//                    worksheet.Cells[1, columnCount++].Value = "Supplier Name";
//                if (Parobj.VFranchise)
//                    worksheet.Cells[1, columnCount++].Value = "Franchise";
//                if (Parobj.VTransactionNumber)
//                    worksheet.Cells[1, columnCount++].Value = "Transaction Number";
//                if (Parobj.VQty)
//                    worksheet.Cells[1, columnCount++].Value = "Total Qty";
//                if (Parobj.VPrice)
//                    worksheet.Cells[1, columnCount++].Value = "Max Price";
//                if (Parobj.VCost)
//                    worksheet.Cells[1, columnCount++].Value = "Cost";
//                if (Parobj.VTotalSales)
//                    worksheet.Cells[1, columnCount++].Value = "Total Sales";
//                if (Parobj.VTransactionCount)
//                    worksheet.Cells[1, columnCount++].Value = "Transactions Count";
//                if (Parobj.VTotalCost)
//                    worksheet.Cells[1, columnCount++].Value = "Total Cost";
//                if (Parobj.VTotalTax)
//                    worksheet.Cells[1, columnCount++].Value = "Tax";
//                if (Parobj.VTotalSalesTax)
//                    worksheet.Cells[1, columnCount++].Value = "Total Sales Tax";
//                if (Parobj.VTotalSalesWithoutTax)
//                    worksheet.Cells[1, columnCount++].Value = "Total Sales Without Tax";
//                if (Parobj.VTotalCostQty)
//                    worksheet.Cells[1, columnCount++].Value = "Total Quantity Cost";
//                // Set header style
//                if (columnCount <= 1)
//                {
//                    // Log a message or throw an exception
//                    Console.WriteLine("Error: columnCount is 0. No data to process.");
//                    // Optionally, throw an exception to halt execution
//                    // throw new InvalidOperationException("columnCount is 0. No data to process.");
//                }
//                else
//                {
//                    using (var headerRange = worksheet.Cells[1, 1, 1, columnCount - 1])
//                    {
//                        headerRange.Style.Font.Bold = true;

//                        // Apply the border style
//                        headerRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
//                        headerRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
//                        headerRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
//                        headerRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
//                        // Apply the horizontal alignment
//                        headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
//                        // Apply the background color
//                        headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
//                        headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SkyBlue);
//                        worksheet.Cells[1, 1, 1, columnCount - 1].AutoFilter = true;
//                    }
//                }
//                int row = 2;
//                foreach (var item in reportData1)
//                {
//                    columnCount = 1; // Reset column count for each row

//                    if (Parobj.VPerYear || Parobj.VPerMonYear)
//                        worksheet.Cells[row, columnCount++].Value = item.PerYear;
//                    if (Parobj.VPerMon || Parobj.VPerMonYear)
//                        worksheet.Cells[row, columnCount++].Value = item.PerMonth;
//                    if (Parobj.VPerDay)
//                        worksheet.Cells[row, columnCount++].Value = item.PerDay;
//                    if (Parobj.VStoreId)
//                        worksheet.Cells[row, columnCount++].Value = item.StoreId;
//                    if (Parobj.VStoreName)
//                        worksheet.Cells[row, columnCount++].Value = item.StoreName;
//                    if (Parobj.VDpId)
//                        worksheet.Cells[row, columnCount++].Value = item.DpId;
//                    if (Parobj.VDepartment)
//                        worksheet.Cells[row, columnCount++].Value = item.DpName;
//                    if (Parobj.VItemLookupCode)
//                        worksheet.Cells[row, columnCount++].Value = item.ItemLookupCodeTxt;
//                    if (Parobj.VItemName)
//                        worksheet.Cells[row, columnCount++].Value = item.ItemName;
//                    if (Parobj.VSupplierId)
//                        worksheet.Cells[row, columnCount++].Value = item.SupplierId;
//                    if (Parobj.VSupplierName)
//                        worksheet.Cells[row, columnCount++].Value = item.SupplierName;
//                    if (Parobj.VFranchise)
//                        worksheet.Cells[row, columnCount++].Value = item.StoreFranchise;
//                    if (Parobj.VTransactionNumber)
//                        worksheet.Cells[row, columnCount++].Value = item.TransactionNumber;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VQty)
//                        worksheet.Cells[row, columnCount++].Value = item.TotalQty;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VPrice)
//                        worksheet.Cells[row, columnCount++].Value = item.Price;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VCost)
//                        worksheet.Cells[row, columnCount++].Value = item.Cost;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VTotalSales)
//                        worksheet.Cells[row, columnCount++].Value = item.Total;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VTransactionCount)
//                        worksheet.Cells[row, columnCount++].Value = item.TransactionCount;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VTotalCost)
//                        worksheet.Cells[row, columnCount++].Value = item.TotalCost;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VTotalTax)
//                        worksheet.Cells[row, columnCount++].Value = item.TotalTax;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VTotalSalesTax)
//                        worksheet.Cells[row, columnCount++].Value = item.TotalSalesTax;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VTotalSalesWithoutTax)
//                        worksheet.Cells[row, columnCount++].Value = item.TotalSalesWithoutTax;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (Parobj.VTotalCostQty)
//                        worksheet.Cells[row, columnCount++].Value = item.TotalCostQty;
//                    worksheet.Cells[row, columnCount].Style.Numberformat.Format = "#,##0.00";
//                    if (columnCount <= 1)
//                    {
//                        Console.WriteLine("Error: columnCount is 0. No data to process.");
//                    }

//                    // Auto fit columns

//                    //    // Log a message or throw an exception
//                    //    Console.WriteLine("Error: columnCount is 0. No data to process.");
//                    //    // Optionally, throw an exception to halt execution
//                    //    // throw new InvalidOperationException("columnCount is 0. No data to process.");
//                    //}
//                }
//                // Save the file
//                var stream = new MemoryStream();
//                package.SaveAs(stream);
//                HttpContext.Session.SetString("ExportStatus", "complete");
//                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "AKSalesReport.xlsx");
//            }
//        }
//        [HttpGet]
//        [ResponseCache(Location = ResponseCacheLocation.None, NoStore = true)]
//        public async Task<IActionResult> LogOut()
//        {
//            // Sign out the user
//            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);

//            // Set a TempData variable to indicate logout
//            TempData["IsLoggedOut"] = true;

//            // Clear session on logout
//            HttpContext.Session.Clear();

//            // Prevent caching by setting appropriate HTTP headers
//            //Response.Headers.Add("Cache-Control", "no-cache, no-store, must-revalidate");
//            //Response.Headers.Add("Pragma", "no-cache");
//            //Response.Headers.Add("Expires", "0");
//            try
//            {
//                if (!Response.Headers.ContainsKey("Cache-Control"))
//                {
//                    Response.Headers.Add("Cache-Control", "no-cache, no-store, must-revalidate");
//                }

//                if (!Response.Headers.ContainsKey("Pragma"))
//                {
//                    Response.Headers.Add("Pragma", "no-cache");
//                }

//                if (!Response.Headers.ContainsKey("Expires"))
//                {
//                    Response.Headers.Add("Expires", "0");
//                }

//                return RedirectToAction("Login", "Login");
//            }

//            catch (Exception ex)
//            {
//                Console.WriteLine($"Exception in LogOut action: {ex.Message}");
//                return RedirectToAction("Login", "Login");
//            }
//        }
//        public IActionResult Privacy()
//        {
//            return View();
//        }

//        public IActionResult index1()
//        {
//            return View();
//        }
//        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
//        public IActionResult Error()
//        {
//            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
//        }
//    }

//}
