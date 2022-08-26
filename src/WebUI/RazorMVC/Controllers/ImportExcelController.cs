using ClosedXML.Excel;
using Dapper;
using DHAAccounts.Models;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;


namespace Accounting.Controllers
{
    public class ImportExcelController : Controller
    {
        private AppDbContext _dbcontext = new AppDbContext();
        //// private AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());
        ApplicationUser Profile = ApplicationUser.GetUserProfile();
        //const string pGLAccountType = "22HOLDENTRY";
        public ActionResult Index()
        {
            return View();
        }

        public static string RemoveSpecialCharacters(string str)
        {
            return Regex.Replace(str, "[^a-zA-Z0-9_. ]+", "", RegexOptions.Compiled);
        }

        #region Import Deposit File
        public ActionResult DepositFile()
        {
            DBUtils.ExecuteSQL("Delete from FileEntriesTemp where (cpUserRefNo = '') or (cpUserRefNo is null) or (cpUserRefNo='0')");
            ImportDepositFileViewModel postingViewModel = new ImportDepositFileViewModel();
            Guid guid = Guid.NewGuid();
            postingViewModel.cpUserRefNo = guid.ToString();
            if (string.IsNullOrEmpty(SysPrefs.DefaultCurrency))
            {
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Unable to find default curreny. Please update preferences.");
            }
            if (string.IsNullOrEmpty(SysPrefs.PostingDate))
            {
                postingViewModel.ErrMessage += Common.GetAlertMessage(1, "Unable to find posting date. Please update preferences.");
            }
            return View("~/Views/ImportExcel/DepositFile.cshtml", postingViewModel);
        }
        [HttpPost]
        public ActionResult DepositFileUploaded(ImportDepositFileViewModel pImportDepositFileViewModel, HttpPostedFileBase file)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            IEnumerable<Sheets> pcheckListViewModel = new List<Sheets>();
            var fileName = ""; var path = "";
            pImportDepositFileViewModel.isFileImported = false;
            try
            {
                if (file != null && file.ContentLength > 0)
                {
                    var supportedTypes = new[] { "xlsx" };
                    var fileExt = Path.GetExtension(file.FileName).Substring(1);
                    if (!supportedTypes.Contains(fileExt))
                    {
                        pImportDepositFileViewModel.ErrMessage = Common.GetAlertMessage(1, "File extension is invalid - Only xlsx are allowed</br>");
                        pImportDepositFileViewModel.isError = true;
                    }
                    else
                    {
                        fileName = Path.GetFileName(file.FileName);
                        ////fileName = Common.getRandomDigit(6) + "-" + fileName;
                        fileName = "Deposit-" + DateTime.Now.ToString("MMddyyyyHHmmssffff") + ".xlsx";
                        path = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/UpTemp/"), fileName);
                        file.SaveAs(path);
                        pImportDepositFileViewModel.ExcelFilePath = path;
                        if (path != null && path != "")
                        {
                            var sheetname1 = "";
                            var workbook = new XLWorkbook(path);
                            if (workbook.Worksheets.Count > 0)
                            {
                                int SheetCount = 1;
                                foreach (var worksheet in workbook.Worksheets)
                                {
                                    if (SheetCount == 1)
                                    {
                                        var sheetname = worksheet.ToString().Split('!');
                                        sheetname1 = sheetname[0].ToString();
                                        sheetname1 = RemoveSpecialCharacters(sheetname1);
                                    }
                                    SheetCount++;
                                }

                                #region import data into temp table

                                IXLWorksheet ws;
                                bool IsRowValid = true;
                                string rowErrMessage = "";
                                string Currency = "";
                                decimal Amount = 0.0m;
                                decimal FcAmount = 0.0m;
                                decimal LcAmount = 0.0m;

                                workbook.TryGetWorksheet(Common.toString(sheetname1), out ws);
                                var datarange = ws.RangeUsed();
                                int TotalRowsinExcelSheet = datarange.RowCount();
                                int TotalColinExcelSheet = datarange.ColumnCount();

                                int TotalFileColumns = 10;
                                //if (SysPrefs.ImportFileHasPostingDate == "1")
                                //{
                                //    TotalFileColumns = 11;
                                //}
                                if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                                {
                                    TotalFileColumns = 11;
                                }
                                if (TotalColinExcelSheet == TotalFileColumns)
                                {
                                    int row_number = 2;
                                    //IEnumerable<ImportDepositFileViewModel> plist = new List<ImportDepositFileViewModel>();
                                    SqlCommand InsertCommand = null;
                                    SqlConnection con = Common.getConnection();
                                    for (int r = row_number; r <= TotalRowsinExcelSheet; r++)
                                    {
                                        string DefaultCurrency = SysPrefs.DefaultCurrency;
                                        string CreditAccountCurrency = "";

                                        string CreditAccountCode = "";
                                        string DebitAccountCurrency = "";
                                        string DebitAccountCode = "";

                                        string sqlMyEntry = "Insert into DepositFileEntriesTemp ";
                                        sqlMyEntry += " (RowNumber, AgentPrefix, CustomerId, CustomerName, Currency, Rate, Amount, Type, Details, Account, FCAmount, LCAmount, isRowValid, ErrorMessage, cpUserRefNo,PostingDate) Values ";
                                        sqlMyEntry += "(@RowNumber, @AgentPrefix, @CustomerId, @CustomerName, @Currency, @Rate, @Amount, @Type, @Details, @Account, @FCAmount, @LCAmount, @isRowValid,@ErrorMessage,@cpUserRefNo,@PostingDate)";

                                        InsertCommand = con.CreateCommand();
                                        if (Common.GetSysSettings("ImportFileHasPostingDate") == "0")
                                        {
                                            if (Common.ValidateFiscalYearDate(Common.toDateTime(SysPrefs.PostingDate)))
                                            {
                                                if (Common.ValidatePostingDate(SysPrefs.PostingDate))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));

                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                    IsRowValid = false;
                                                    rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                                }
                                            }
                                            else
                                            {
                                                InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                IsRowValid = false;
                                                rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                                            }
                                        }
                                        for (int i = 1; i <= TotalColinExcelSheet; i++)
                                        {
                                            //ImportDepositFileViewModel list = new ImportDepositFileViewModel();
                                            if (i == 1) // AgentPrefix
                                            {
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@AgentPrefix", ws.Row(r).Cell(i).Value);

                                                    DataTable dtCredit = DBUtils.GetDataTable("select AccountCode, CurrencyISOCode from GLChartOfAccounts where prefix = '" + Common.toString(ws.Row(r).Cell(i).Value) + "'");
                                                    if (dtCredit != null && dtCredit.Rows.Count > 0)
                                                    {
                                                        CreditAccountCode = Common.toString(dtCredit.Rows[0]["AccountCode"]);
                                                        CreditAccountCurrency = Common.toString(dtCredit.Rows[0]["CurrencyISOCode"]);
                                                    }
                                                    if (CreditAccountCode == "")
                                                    {
                                                        IsRowValid = false;
                                                        rowErrMessage += "Invalid Agent Prefix.  Agent prefix not available in chart of account.</br>";
                                                    }
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@AgentPrefix", "");
                                                    IsRowValid = false;
                                                    rowErrMessage += "Agent Prefix is not available in excel sheet.</br>";
                                                }
                                            }

                                            if (i == 2) // CustomerId
                                            {
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@CustomerId", ws.Row(r).Cell(i).Value.ToString());
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@CustomerId", "");
                                                }
                                            }

                                            if (i == 3) // CustomerName
                                            {
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@CustomerName", ws.Row(r).Cell(i).Value.ToString());
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@CustomerName", "");
                                                }
                                            }

                                            if (i == 4) // Currency
                                            {
                                                Currency = Common.toString(ws.Row(r).Cell(i).Value);
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@Currency", ws.Row(r).Cell(i).Value.ToString());

                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@Currency", "");
                                                    IsRowValid = false;
                                                    rowErrMessage += "Currency is not available in excel sheet<br>";
                                                }

                                            }

                                            if (i == 5) // Amount
                                            {
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    Amount = Common.toDecimal(ws.Row(r).Cell(i).Value.ToString());
                                                    InsertCommand.Parameters.AddWithValue("@Amount", Amount);
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@Amount", 0);
                                                    IsRowValid = false;
                                                    rowErrMessage += "Amount is not available in excel sheet<br>";
                                                }
                                            }

                                            if (i == 6) // Type
                                            {
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@Type", ws.Row(r).Cell(i).Value.ToString());
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@Type", "");
                                                }
                                            }

                                            if (i == 7) // Details
                                            {
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@Details", ws.Row(r).Cell(i).Value.ToString());
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@Details", "");
                                                    IsRowValid = false;
                                                    rowErrMessage += "Details are not available in excel sheet<br>";
                                                }
                                            }

                                            if (i == 8) // Account
                                            {
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@Account", ws.Row(r).Cell(i).Value.ToString());
                                                    DataTable dtDedit = DBUtils.GetDataTable("select AccountCode, CurrencyISOCode from GLChartOfAccounts where prefix = '" + Common.toString(ws.Row(r).Cell(i).Value) + "'");
                                                    if (dtDedit != null && dtDedit.Rows.Count > 0)
                                                    {
                                                        DebitAccountCode = Common.toString(dtDedit.Rows[0]["AccountCode"]);
                                                        DebitAccountCurrency = Common.toString(dtDedit.Rows[0]["CurrencyISOCode"]);
                                                    }
                                                    if (DebitAccountCode == "")
                                                    {
                                                        IsRowValid = false;
                                                        rowErrMessage += " Debit [Account] is not available in chart of account<br>";
                                                    }
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@Account", "");
                                                    IsRowValid = false;
                                                    rowErrMessage += "Account is not available in excel sheet.<br>";
                                                }
                                            }

                                            if (i == 9) // FCAmount
                                            {
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    FcAmount = Common.toDecimal(ws.Row(r).Cell(i).Value.ToString());
                                                    InsertCommand.Parameters.AddWithValue("@FCAmount", FcAmount);
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@FCAmount", 0);
                                                    IsRowValid = false;
                                                    rowErrMessage += "L/C Amount is not available in excel sheet<br>";
                                                }
                                            }

                                            if (i == 10) // LCAmount
                                            {
                                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                {
                                                    LcAmount = Common.toDecimal(ws.Row(r).Cell(i).Value.ToString());
                                                    InsertCommand.Parameters.AddWithValue("@LCAmount", LcAmount);
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@LCAmount", 0);
                                                    IsRowValid = false;
                                                    rowErrMessage += "L/C Amount is not available in excel sheet<br>";
                                                }
                                            }

                                            if (i == 11) // Posting Date
                                            {
                                                if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                                                {
                                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                                    {
                                                        //is valid date
                                                        //is within fin year
                                                        string pDate = Common.toString(ws.Row(r).Cell(i).Value);
                                                        if (Common.isDate(pDate))
                                                        {
                                                            if (Common.ValidateFiscalYearDate(Common.toDateTime(ws.Row(r).Cell(i).Value)))
                                                            {
                                                                if (Common.ValidatePostingDate(Common.toString(ws.Row(r).Cell(i).Value)))
                                                                {
                                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(ws.Row(r).Cell(i).Value));
                                                                }
                                                                else
                                                                {
                                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                                    IsRowValid = false;
                                                                    rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                                                }

                                                            }
                                                            else
                                                            {
                                                                InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                                IsRowValid = false;
                                                                rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                                                            }

                                                        }
                                                        else
                                                        {
                                                            InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                            IsRowValid = false;
                                                            rowErrMessage += "Posting date is not valid.<br>";
                                                        }

                                                    }
                                                    else
                                                    {
                                                        InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                        IsRowValid = false;
                                                        rowErrMessage += "Posting date is not available in excel sheet.<br>";
                                                    }

                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                                                }
                                            }



                                            #region attachment code
                                            //if (i == 11) // Attachment1
                                            //{
                                            //    string name = "";
                                            //    string FileName = "";
                                            //    bool ifExist = false;
                                            //    string extension = "";
                                            //    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                            //    {
                                            //        string Lpath = Common.toString(ws.Row(r).Cell(i).Value).Trim();
                                            //        string[] getExtension = Lpath.Split(Path.DirectorySeparatorChar);
                                            //        string getlast = getExtension.Last();
                                            //        string[] ext = getlast.Split('.');
                                            //        extension = ext.Last();
                                            //        string currentDirectory = Path.GetDirectoryName(Lpath);
                                            //        string BaseFolder = Path.GetFullPath(currentDirectory);
                                            //        foreach (string txtName in Directory.GetFiles(BaseFolder, "*." + extension + ""))
                                            //        {
                                            //            if (txtName == Lpath)
                                            //            {
                                            //                FileName = Path.GetFileName(txtName);
                                            //                FileName = pImportDepositFileViewModel.cpUserRefNo + "-" + Common.getRandomDigit(6) + "-" + FileName;
                                            //                Uri uri = new Uri(txtName.Replace("/download", ""));
                                            //                string myfilename = System.IO.Path.GetFileName(uri.LocalPath);
                                            //                string savePath = Path.Combine(Server.MapPath("~/Uploads/Vouchers/"));
                                            //                if (!System.IO.Directory.Exists(savePath)) System.IO.Directory.CreateDirectory(savePath);
                                            //                savePath += "/" + FileName;
                                            //                WebClient _webclient = new WebClient();
                                            //                _webclient.DownloadFile(txtName, savePath);
                                            //                ifExist = true;
                                            //            }
                                            //        }

                                            //        extension = "." + extension;
                                            //        string Fpath = ext[0].ToString();
                                            //        string[] getName = Fpath.Split('\\');
                                            //        name = getName.Last();
                                            //        list.Extension = extension;
                                            //        list.Name = name + extension;
                                            //        list.Size = 5000;
                                            //        plist = (new[] { list }).Concat(plist);
                                            //        if (ifExist)
                                            //        {
                                            //            InsertCommand.Parameters.AddWithValue("@File1", Common.toString(FileName));
                                            //        }
                                            //        else
                                            //        {
                                            //            InsertCommand.Parameters.AddWithValue("@File1", "");
                                            //            IsRowValid = false;
                                            //            rowErrMessage += "" + name + " is not available in the given location<br>";
                                            //        }
                                            //    }
                                            //    else
                                            //    {
                                            //        InsertCommand.Parameters.AddWithValue("@File1", "");
                                            //    }
                                            //}

                                            //if (i == 12) // Attachment2
                                            //{
                                            //    string name = "";
                                            //    string FileName = "";
                                            //    bool ifExist = false;
                                            //    string extension = "";
                                            //    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                            //    {
                                            //        string Lpath = Common.toString(ws.Row(r).Cell(i).Value).Trim();
                                            //        string[] getExtension = Lpath.Split(Path.DirectorySeparatorChar);
                                            //        string getlast = getExtension.Last();
                                            //        string[] ext = getlast.Split('.');
                                            //        extension = ext.Last();
                                            //        string currentDirectory = Path.GetDirectoryName(Lpath);
                                            //        string BaseFolder = Path.GetFullPath(currentDirectory);
                                            //        foreach (string txtName in Directory.GetFiles(BaseFolder, "*." + extension + ""))
                                            //        {
                                            //            if (txtName == Lpath)
                                            //            {
                                            //                FileName = Path.GetFileName(txtName);
                                            //                FileName = pImportDepositFileViewModel.cpUserRefNo + "-" + Common.getRandomDigit(6) + "-" + FileName;
                                            //                Uri uri = new Uri(txtName.Replace("/download", ""));
                                            //                string myfilename = System.IO.Path.GetFileName(uri.LocalPath);
                                            //                string savePath = Path.Combine(Server.MapPath("~/Uploads/Vouchers/"));
                                            //                //string savePath = @"C:/inetpub/wwwroot/iRemitifyAccounts/DHAAccounts/Uploads/Vouchers";
                                            //                if (!System.IO.Directory.Exists(savePath)) System.IO.Directory.CreateDirectory(savePath);
                                            //                savePath += "/" + FileName;
                                            //                WebClient _webclient = new WebClient();
                                            //                _webclient.DownloadFile(txtName, savePath);
                                            //                ifExist = true;
                                            //            }
                                            //        }

                                            //        extension = "." + extension;
                                            //        string Fpath = ext[0].ToString();
                                            //        string[] getName = Fpath.Split('\\');
                                            //        name = getName.Last();
                                            //        list.Extension = extension;
                                            //        list.Name = name + extension;
                                            //        list.Size = 5000;
                                            //        plist = (new[] { list }).Concat(plist);
                                            //        if (ifExist)
                                            //        {
                                            //            InsertCommand.Parameters.AddWithValue("@File2", Common.toString(FileName));
                                            //        }
                                            //        else
                                            //        {
                                            //            InsertCommand.Parameters.AddWithValue("@File2", "");
                                            //            IsRowValid = false;
                                            //            rowErrMessage += "" + name + " is not available in the given location<br>";
                                            //        }
                                            //    }
                                            //    else
                                            //    {
                                            //        InsertCommand.Parameters.AddWithValue("@File2", "");
                                            //    }
                                            //}
                                            //if (i == 13) // Attachment3
                                            //{
                                            //    string name = "";
                                            //    string FileName = "";
                                            //    bool ifExist = false;
                                            //    string extension = "";
                                            //    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                            //    {
                                            //        string Lpath = Common.toString(ws.Row(r).Cell(i).Value).Trim();
                                            //        string[] getExtension = Lpath.Split(Path.DirectorySeparatorChar);
                                            //        string getlast = getExtension.Last();
                                            //        string[] ext = getlast.Split('.');
                                            //        extension = ext.Last();
                                            //        string currentDirectory = Path.GetDirectoryName(Lpath);
                                            //        string BaseFolder = Path.GetFullPath(currentDirectory);
                                            //        foreach (string txtName in Directory.GetFiles(BaseFolder, "*." + extension + ""))
                                            //        {
                                            //            if (txtName == Lpath)
                                            //            {
                                            //                FileName = Path.GetFileName(txtName);
                                            //                FileName = pImportDepositFileViewModel.cpUserRefNo + "-" + Common.getRandomDigit(6) + "-" + FileName;
                                            //                Uri uri = new Uri(txtName.Replace("/download", ""));
                                            //                string myfilename = System.IO.Path.GetFileName(uri.LocalPath);
                                            //                string savePath = Path.Combine(Server.MapPath("~/Uploads/Vouchers/"));
                                            //                if (!System.IO.Directory.Exists(savePath)) System.IO.Directory.CreateDirectory(savePath);
                                            //                savePath += "/" + FileName;
                                            //                WebClient _webclient = new WebClient();
                                            //                _webclient.DownloadFile(txtName, savePath);
                                            //                ifExist = true;
                                            //            }
                                            //        }

                                            //        extension = "." + extension;
                                            //        string Fpath = ext[0].ToString();
                                            //        string[] getName = Fpath.Split('\\');
                                            //        name = getName.Last();
                                            //        list.Extension = extension;
                                            //        list.Name = name + extension;
                                            //        list.Size = 5000;
                                            //        plist = (new[] { list }).Concat(plist);
                                            //        if (ifExist)
                                            //        {
                                            //            InsertCommand.Parameters.AddWithValue("@File3", Common.toString(FileName));
                                            //        }
                                            //        else
                                            //        {
                                            //            InsertCommand.Parameters.AddWithValue("@File3", "");
                                            //            IsRowValid = false;
                                            //            rowErrMessage += "" + name + " is not available in the given location<br>";
                                            //        }
                                            //    }
                                            //    else
                                            //    {
                                            //        InsertCommand.Parameters.AddWithValue("@File3", "");
                                            //    }
                                            //}
                                            #endregion
                                        }
                                        // Rate calculation
                                        if (Common.toString(DebitAccountCurrency).Trim() != "" && Common.toString(CreditAccountCurrency).Trim() != "")
                                        {
                                            decimal CreditMinRate = 0m;
                                            decimal CreditMaxRate = 0m;
                                            decimal DebitMinRate = 0m;
                                            decimal DebitMaxRate = 0m;
                                            decimal DebitExchangeRate = 0m;
                                            decimal CreditExchangeRate = 0m;

                                            DataTable objCreditCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + CreditAccountCurrency + "'");
                                            if (objCreditCurrency != null)
                                            {
                                                if (objCreditCurrency.Rows.Count > 0)
                                                {
                                                    DataRow drCreditCurrency = objCreditCurrency.Rows[0];
                                                    CreditMinRate = Common.toDecimal(drCreditCurrency["MinRateLimit"]);
                                                    CreditMaxRate = Common.toDecimal(drCreditCurrency["MaxRateLimit"]);
                                                }
                                            }

                                            DataTable objDebitCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + DebitAccountCurrency + "'");
                                            if (objDebitCurrency != null)
                                            {
                                                if (objDebitCurrency.Rows.Count > 0)
                                                {
                                                    DataRow drDebitCurrency = objDebitCurrency.Rows[0];
                                                    DebitMinRate = Common.toDecimal(drDebitCurrency["MinRateLimit"]);
                                                    DebitMaxRate = Common.toDecimal(drDebitCurrency["MaxRateLimit"]);
                                                }
                                            }

                                            if (DebitAccountCurrency == DefaultCurrency && CreditAccountCurrency == DefaultCurrency)
                                            {
                                                CreditExchangeRate = 1;
                                                DebitExchangeRate = 1;
                                            }
                                            else if (DebitAccountCurrency != DefaultCurrency && CreditAccountCurrency == DefaultCurrency)
                                            {
                                                CreditExchangeRate = 1;
                                                if (Amount > 0)
                                                {
                                                    DebitExchangeRate = FcAmount / Amount;
                                                }
                                            }
                                            else if (DebitAccountCurrency == DefaultCurrency && CreditAccountCurrency != DefaultCurrency)
                                            {
                                                if (LcAmount > 0)
                                                {
                                                    CreditExchangeRate = Amount / LcAmount;
                                                }
                                                DebitExchangeRate = 1;

                                            }
                                            else if (DebitAccountCurrency != DefaultCurrency && CreditAccountCurrency != DefaultCurrency)
                                            {
                                                //CreditExchangeRate = Amount / LcAmount;
                                                if (LcAmount > 0)
                                                {
                                                    CreditExchangeRate = FcAmount / LcAmount;
                                                    DebitExchangeRate = FcAmount / LcAmount;
                                                }
                                            }

                                            if (CreditExchangeRate > 0 && DebitExchangeRate > 0)
                                            {
                                                InsertCommand.Parameters.AddWithValue("@Rate", CreditExchangeRate);

                                                if (CreditMinRate > 0 && CreditMaxRate > 0)
                                                {
                                                    if (CreditMinRate > CreditExchangeRate || CreditMaxRate < CreditExchangeRate)
                                                    {
                                                        IsRowValid = false;
                                                        rowErrMessage += "Credit Rate can not be greater then max rate [" + CreditMaxRate.ToString() + "] or less then min rate [" + CreditMinRate.ToString() + "]<br/>";
                                                    }
                                                }
                                                if (DebitMinRate > 0 && DebitMaxRate > 0)
                                                {
                                                    if (DebitMinRate > DebitExchangeRate || DebitMaxRate < DebitExchangeRate)
                                                    {
                                                        IsRowValid = false;
                                                        rowErrMessage += "Debit Rate can not be greater then max rate [" + DebitMaxRate.ToString() + "] or less then min rate [" + DebitMinRate.ToString() + "]<br/>";
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                InsertCommand.Parameters.AddWithValue("@Rate", 0);
                                            }
                                        }

                                        else
                                        {
                                            InsertCommand.Parameters.AddWithValue("@Rate", 0);
                                        }

                                        if (IsRowValid)
                                        {
                                            InsertCommand.Parameters.AddWithValue("@isRowValid", true);
                                            InsertCommand.Parameters.AddWithValue("@cpUserRefNo", pImportDepositFileViewModel.cpUserRefNo);
                                            InsertCommand.Parameters.AddWithValue("@ErrorMessage", rowErrMessage);
                                        }
                                        else
                                        {
                                            InsertCommand.Parameters.AddWithValue("@isRowValid", false);
                                            InsertCommand.Parameters.AddWithValue("@ErrorMessage", rowErrMessage);
                                            InsertCommand.Parameters.AddWithValue("@cpUserRefNo", pImportDepositFileViewModel.cpUserRefNo);
                                        }

                                        InsertCommand.Parameters.AddWithValue("@RowNumber", r - 1);

                                        InsertCommand.CommandText = sqlMyEntry;
                                        InsertCommand.ExecuteNonQuery();
                                        InsertCommand.Parameters.Clear();
                                        rowErrMessage = "";
                                        IsRowValid = true;
                                        pImportDepositFileViewModel.isPosted = true;
                                    }
                                    //pImportDepositFileViewModel.uploadFilesList = plist.ToList();
                                    con.Close();
                                }
                                else
                                {
                                    pImportDepositFileViewModel.ErrMessage = "Incomplete data in excel sheet. Please check total number of columns in excel sheet";
                                }
                                #endregion

                                pImportDepositFileViewModel.isFileImported = true;
                            }
                        }
                        else
                        {
                            pImportDepositFileViewModel.ErrMessage = "file not uploaded. please try again.";
                            pImportDepositFileViewModel.isError = true;
                        }
                    }
                }
                else
                {
                    pImportDepositFileViewModel.ErrMessage = "file not uploaded. please try again.";
                    pImportDepositFileViewModel.isError = true;
                }
            }
            catch (Exception ex)
            {
                pImportDepositFileViewModel.isFileImported = false;
                pImportDepositFileViewModel.ErrMessage = "Unable to load data Please try Again -> " + ex.ToString();
            }

            return View("~/Views/ImportExcel/DepositFile.cshtml", pImportDepositFileViewModel);
        }

        public ActionResult ImportDepositCheckList_Read([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM DepositFileEntriesTemp /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                          Id,RowNumber,cpUserRefNo, AgentPrefix, CustomerId, PostingDate,CustomerName,Currency,Amount,Type,Details,Account,FCAmount, LCAmount, Rate, isRowValid,ErrorMessage FROM DepositFileEntriesTemp
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";
            //     const string selectQuery = @"SELECT
            //CheckListId, CheckTypeId, (select CheckType  from CheckTypes where CheckTypes.CheckTypeId = CheckList.CheckTypeId ) as CheckTypeName, CheckText, case when Status = 'true' then 'Active' else 'Inactive' end as StatusName, Status FROM CheckList";
            SqlBuilder builder = new SqlBuilder();
            var count = builder.AddTemplate(countQuery);

            int CurrentPage = 0; int CurrentPageSize = 1;
            if (request.PageSize > 0)
            {
                CurrentPageSize = request.PageSize;
            }
            if (request.PageSize > 0 && request.Page > 0)
            {
                CurrentPage = request.Page - 1;
            }
            var selector = builder.AddTemplate(selectQuery, new { PageIndex = CurrentPage, PageSize = CurrentPageSize });

            if (request.Filters != null && request.Filters.Any())
            {
                builder = Common.ApplyFilters(builder, request.Filters);
            }
            else
            {
                builder.Where("cpUserRefNo=''");
            }

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("RowNumber");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);

            var rows = _dbcontext.Query<ImportDepositFileViewModel>(selector.RawSql, selector.Parameters);

            //  rows.Each(item => Console.WriteLine($"({ item.AccountCode}): { item.AccountName}
            //{                GetParentsString(items, item)}"));            
            //}
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult DepositEditingInline_Update([DataSourceRequest] DataSourceRequest request, ImportDepositFileViewModel product)
        {
            if (product != null)
            {
                product = DepositEditingCustom_Update(product);
            }

            return Json(new[] { product }.ToDataSourceResult(request, ModelState));
        }

        public ImportDepositFileViewModel DepositEditingCustom_Update(ImportDepositFileViewModel pFileModel)
        {
            ApplicationUser Profile = new ApplicationUser();
            bool IsRowValid = true;
            string err_msgs = "";
            string UpdateSQL = "Update DepositFileEntriesTemp Set ";
            UpdateSQL += "AgentPrefix=@AgentPrefix, CustomerId=@CustomerId,PostingDate=@PostingDate, CustomerName=@CustomerName, Currency=@Currency, Amount=@Amount, Type=@Type, Details=@Details, Account=@Account, FCAmount=@FCAmount, LCAmount=@LCAmount, isRowValid=@isRowValid, ErrorMessage=@ErrorMessage Where Id=@Id";
            using (SqlConnection con = Common.getConnection())
            {
                SqlCommand updateCommand = con.CreateCommand();
                try
                {
                    if (Common.toString(pFileModel.AgentPrefix).Trim() != "")
                    {
                        updateCommand.Parameters.AddWithValue("@AgentPrefix", pFileModel.AgentPrefix);
                        string accountcode = Common.getGLAccountCodeByPrefix(pFileModel.AgentPrefix);
                        if (accountcode == "")
                        {
                            IsRowValid = false;
                            err_msgs += "Invalid Agent prefix. Agent Prefix not available in chart of account;";
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Agent Prefix is not available in editable grid; ";
                        updateCommand.Parameters.AddWithValue("@AgentPrefix", "");
                    }

                    // Customer Id
                    if (pFileModel.CustomerId > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerId", pFileModel.CustomerId);
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerId", "");
                    }

                    if (Common.toString(pFileModel.PostingDate).Trim() != "")
                    {
                        if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                        {
                            if (Common.isDate(Common.toString(pFileModel.PostingDate)))
                            {
                                if (Common.ValidateFiscalYearDate(Common.toDateTime(pFileModel.PostingDate)))
                                {
                                    if (Common.ValidatePostingDate(Common.toString(pFileModel.PostingDate)))
                                    {
                                        updateCommand.Parameters.AddWithValue("@PostingDate", pFileModel.PostingDate);
                                    }
                                    else
                                    {
                                        updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                        IsRowValid = false;
                                        err_msgs += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                    }
                                }
                                else
                                {
                                    updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                    IsRowValid = false;
                                    err_msgs += "Posting date is not between financial year start date and financial year end date.<br>";
                                }

                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Posting date is not valid.<br>";
                            }

                        }
                        else
                        {
                            if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", SysPrefs.PostingDate);
                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                            }

                        }

                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Posting date is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@PostingDate", "");
                    }

                    // Customer Name
                    if (!string.IsNullOrEmpty(pFileModel.CustomerName))
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerName", pFileModel.CustomerName);
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerName", "");
                    }

                    // PayOut_CCY
                    if (!string.IsNullOrEmpty(pFileModel.Account))
                    {
                        string accountcode = Common.getGLAccountCodeByPrefix(pFileModel.Account);
                        if (accountcode == "")
                        {
                            IsRowValid = false;
                            err_msgs += pFileModel.Account + " Debit [Account] is not available in chart of account;";
                        }
                        updateCommand.Parameters.AddWithValue("@Account", pFileModel.Account);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Debit [Account] is not available in editable grid;";
                        updateCommand.Parameters.AddWithValue("@Account", "");
                    }

                    // Amount
                    if (pFileModel.Amount >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Amount", pFileModel.Amount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Amount not available in editable grid;";
                        updateCommand.Parameters.AddWithValue("@Amount", 0);
                    }

                    // Type
                    if (!string.IsNullOrEmpty(pFileModel.Type))
                    {
                        updateCommand.Parameters.AddWithValue("@Type", pFileModel.Type);
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@Type", "");
                    }

                    /// Details
                    if (!string.IsNullOrEmpty(pFileModel.Details))
                    {
                        updateCommand.Parameters.AddWithValue("@Details", pFileModel.Details);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Details are not available in editable grid;";
                        updateCommand.Parameters.AddWithValue("@Details", "");
                    }

                    //// FC Amount
                    if (pFileModel.LCAmount >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@FCAmount", pFileModel.LCAmount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "F/C Amount is not available in editable grid; ";
                        updateCommand.Parameters.AddWithValue("@FCAmount", 0);
                    }
                    //// LC Amount
                    if (pFileModel.LCAmount >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@LCAmount", pFileModel.LCAmount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "L/C Amount is not available in editable grid; ";
                        updateCommand.Parameters.AddWithValue("@LCAmount", 0);
                    }

                    if (!string.IsNullOrEmpty(pFileModel.Currency))
                    {
                        updateCommand.Parameters.AddWithValue("@Currency", pFileModel.Currency);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Currency is not available; ";
                        updateCommand.Parameters.AddWithValue("@Currency", "");
                    }

                    if (IsRowValid)
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", true);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                        updateCommand.Parameters.AddWithValue("@ErrorMessage", err_msgs);
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", false);
                        updateCommand.Parameters.AddWithValue("@ErrorMessage", err_msgs);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                    }
                    updateCommand.CommandText = UpdateSQL;
                    updateCommand.ExecuteNonQuery();
                    updateCommand.Parameters.Clear();
                }
                catch (Exception ex)
                {
                }
                finally
                {
                    con.Close();
                }
            }
            pFileModel.isRowValid = IsRowValid;
            pFileModel.RowErrMessage = err_msgs;
            return pFileModel;
        }

        public ActionResult PostDepositData(ImportDepositFileViewModel pPostingModel, HttpPostedFileBase file)
        {
            // iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
            HoldFilePostingViewModel pHoldFilePostingModel = new HoldFilePostingViewModel();
            ErrorMessages objErrorMessage = null;
            ////string myConnectionString = BaseModel.getConnString();
            string myConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
            ApplicationUser profile = ApplicationUser.GetUserProfile();
            DataSet DS_lstGLEntriesTemp = DBUtils.GetDataSet("select * from DepositFileEntriesTemp where cpUserRefNo='" + pPostingModel.cpUserRefNo + "'");
            DataTable lstGLEntriesTemp = DS_lstGLEntriesTemp.Tables[0];
            if (lstGLEntriesTemp != null)
            {
                // GLReferences gLReferences = db.GLReferences.Where(x => x.ReferenceNo == pPostingModel.cpUserRefNo).SingleOrDefault();
                // if (gLReferences != null)
                string gLReferences = DBUtils.executeSqlGetSingle("select GLReferenceID from GLReferences where ReferenceNo ='" + pPostingModel.cpUserRefNo + "'");
                if (!string.IsNullOrEmpty(gLReferences))
                {
                    pPostingModel.ErrMessage = Common.GetAlertMessage(1, "File already imported. Please import new file.");
                }
                else
                {
                    if (lstGLEntriesTemp.Rows.Count > 0)
                    {
                        bool isMainValid = true;
                        foreach (DataRow objGLEntriesTemp in lstGLEntriesTemp.Rows)
                        {
                            if (Common.toBool(objGLEntriesTemp["isRowValid"]) == false)
                            {
                                isMainValid = false;
                            }
                        }
                        if (isMainValid)
                        {
                            //string filePath = Server.MapPath("~/Uploads/Vouchers/");

                            objErrorMessage = PostingUtils.SaveDepositTransactionBridge("22BANKPMT", lstGLEntriesTemp, Profile.Id.ToString());
                            string postingErr = "";
                            if (objErrorMessage.IsFailed)
                            {
                                for (int i = 0; i < objErrorMessage.ErrorMessage.Count; i++)
                                {
                                    postingErr = postingErr + "<br/>" + objErrorMessage.ErrorMessage[i].Message;
                                }
                                pPostingModel.ErrMessage = Common.GetAlertMessage(1, postingErr);
                            }
                            else
                            {

                                pPostingModel.ErrMessage = Common.GetAlertMessage(0, "Deposit file (Excel) imported successfully.<br> <br> " + objErrorMessage.ErrorMessage[0].Message + " <br><a href ='/Inquiries/GLInquiry'>Click here</a>");
                            }

                        }
                        else
                        {
                            pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Please check your data and try again");
                        }
                    }
                }
            }
            else
            {
                pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Data not found");
            }
            pHoldFilePostingModel.ErrMessage = pPostingModel.ErrMessage;
            return View("~/Views/ImportExcel/Message.cshtml", pHoldFilePostingModel);
        }

        #endregion

        #region Import Hold File

        public ActionResult HoldFile()
        {
            DBUtils.ExecuteSQL("Delete from FileEntriesTemp where (cpUserRefNo = '') or (cpUserRefNo is null) or (cpUserRefNo='0')");

            HoldFilePostingViewModel postingViewModel = new HoldFilePostingViewModel();
            Guid guid = Guid.NewGuid();
            postingViewModel.cpUserRefNo = guid.ToString();
            if (string.IsNullOrEmpty(SysPrefs.TransactionHoldingAccount))
            {
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Unable to find Transaction hold classification. Please update prefrences.");
            }
            if (string.IsNullOrEmpty(SysPrefs.TransactionAdminCommissionAccount))
            {
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Unable to find admin charge income account. Please update prefrences.");
            }
            if (string.IsNullOrEmpty(SysPrefs.TransactionAgentCommissionAccount))
            {
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Unable to find agent charge income account. Please update prefrences.");
            }
            return View("~/Views/ImportExcel/HoldFile.cshtml", postingViewModel);
        }

        [HttpPost]
        public ActionResult HoldFileUploaded(HoldFilePostingViewModel pHoldFilePostingViewModel, HttpPostedFileBase file)
        {
            try
            {
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                IEnumerable<Sheets> pcheckListViewModel = new List<Sheets>();
                var fileName = ""; var path = "";
                pHoldFilePostingViewModel.isFileImported = false;
                if (file != null && file.ContentLength > 0)
                {
                    var supportedTypes = new[] { "xlsx" };
                    var fileExt = Path.GetExtension(file.FileName).Substring(1);
                    if (!supportedTypes.Contains(fileExt))
                    {
                        pHoldFilePostingViewModel.ErrMessage = Common.GetAlertMessage(1, "File extension is invalid -  only xlsx files are allowed.");
                        pHoldFilePostingViewModel.isError = true;
                    }
                    else
                    {
                        fileName = Path.GetFileName(file.FileName);
                        //// fileName = Common.getRandomDigit(6) + "-" + fileName;
                        //// path = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/"), fileName);
                        fileName = "Hold-" + DateTime.Now.ToString("MMddyyyyHHmmssffff") + ".xlsx";
                        path = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/UpTemp/"), fileName);
                        file.SaveAs(path);
                        pHoldFilePostingViewModel.isFileImported = true;
                        pHoldFilePostingViewModel.ExcelFilePath = path;
                        if (path != null && path != "")
                        {
                            int Count = 0;
                            var workbook = new XLWorkbook(path);
                            foreach (var worksheet in workbook.Worksheets)
                            {
                                var sheetname = worksheet.ToString().Split('!');
                                var sheetname1 = sheetname[0].ToString();
                                sheetname1 = RemoveSpecialCharacters(sheetname1);

                                Sheets pSheets = new Sheets();
                                pSheets.Name = Common.toString(sheetname1);
                                pSheets.Name = pSheets.Name.Replace("'", "");
                                pSheets.Value = Count.ToString();
                                pcheckListViewModel = (new[] { pSheets }).Concat(pcheckListViewModel);
                                Count++;
                            }
                            pHoldFilePostingViewModel.sheets = pcheckListViewModel.ToList();
                        }
                        else
                        {
                            pHoldFilePostingViewModel.ErrMessage = Common.GetAlertMessage(1, "file not uploaded. please try again.");
                            pHoldFilePostingViewModel.isError = true;
                        }
                    }
                }
                else
                {
                    pHoldFilePostingViewModel.ErrMessage = Common.GetAlertMessage(1, "file not uploaded. please try again.");
                    pHoldFilePostingViewModel.isError = true;
                }
            }
            catch (Exception ex)
            {
                pHoldFilePostingViewModel.ErrMessage = "An error occured: " + ex.ToString();
                pHoldFilePostingViewModel.isError = true;
            }
            return View("~/Views/ImportExcel/HoldFile.cshtml", pHoldFilePostingViewModel);
        }
        public ActionResult ShowExcelfiledata(HoldFilePostingViewModel viewModel)
        {
            string sqlMyEntry = "";
            try
            {
                string autoSetPostingDate = Common.GetSysSettings("AutoSetPostingDate");
                string savepath = viewModel.ExcelFilePath;
                var workbook = new XLWorkbook(savepath);

                string PayoutCurrency = "";
                bool IsRowValid = true;
                string rowErrMessage = "";
                string sheetname = viewModel.Sheet;
                IXLWorksheet ws;
                workbook.TryGetWorksheet(Common.toString(sheetname), out ws);
                var datarange = ws.RangeUsed();
                int TotalRowsinExcelSheet = datarange.RowCount();
                int TotalColinExcelSheet = datarange.ColumnCount();

                int TotalFileColumns = 22;
                if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                {
                    TotalFileColumns = 23;
                }
                if (TotalColinExcelSheet == TotalFileColumns)
                {
                    int row_number = 0;
                    if (viewModel.HasHeader)
                    {
                        row_number = 2;
                    }
                    else
                    {
                        row_number = 1;
                    }
                    int row_counter = 0;
                    int j = 1;
                    SqlCommand InsertCommand = null;
                    SqlConnection con = Common.getConnection();
                    for (int r = row_number; r <= TotalRowsinExcelSheet; r++)
                    {
                        sqlMyEntry = "Insert into FileEntriesTemp (RowNumber,TransactionDate,AgentName,Rate,PayoutCurrency,FCAmount,Payinamount, AdminCharges,AgentCharges,TRANID,CustomerId,CustomerName,FatherName,Recipient,Phone,Address,City,Code,PaymentNo,BenePaymentMethod,SendingCountry,ReceivingCountry,Status, isRowValid, RowErrMessage,cpUserRefNo,PostingDate) Values (@RowNumber,@TransactionDate,@AgentName,@Rate,@PayoutCurrency,@FCAmount,@Payinamount, @AdminCharges,@AgentCharges,@TRANID,@CustomerId,@CustomerName,@FatherName,@Recipient,@Phone,@Address,@City,@Code,@PaymentNo,@BenePaymentMethod,@SendingCountry,@ReceivingCountry,@Status,@isRowValid,@RowErrMessage,@cpUserRefNo,@PostingDate)";
                        string AgentName = "";
                        string AgentAccountCode = "";
                        string AgentCurrencyCode = "";
                        InsertCommand = con.CreateCommand();
                        //if (ws.Row(r).Cell(22).Value.ToString() != "Canceled" && ws.Row(r).Cell(22).Value.ToString() != "Refunded")
                        //{
                        if (Common.GetSysSettings("ImportFileHasPostingDate") == "0")
                        {
                            if (Common.ValidateFiscalYearDate(Common.toDateTime(SysPrefs.PostingDate)))
                            {
                                if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                                {
                                    InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                }
                            }
                            else
                            {
                                InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                            }
                        }
                        for (int i = 1; i <= TotalColinExcelSheet; i++)
                        {
                            if (i == 1) // Agent Name
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    AgentName = Common.toString(ws.Row(r).Cell(i).Value);
                                    InsertCommand.Parameters.AddWithValue("@AgentName", AgentName);
                                    AgentAccountCode = Common.getGLAccountCodeByPrefix(AgentName);
                                    if (AgentAccountCode.Trim() != "")
                                    {
                                        AgentCurrencyCode = Common.getGLAccountCurrencyByCode(AgentAccountCode);
                                        if (AgentCurrencyCode != "")
                                        {
                                            if (!Common.isCurrencyAvailable(AgentCurrencyCode))
                                            {
                                                IsRowValid = false;
                                                rowErrMessage += AgentCurrencyCode + " is not available or is inactive. Please check currency settings<br>";
                                                AgentCurrencyCode = "";
                                            }
                                            else
                                            {
                                                string sqlCustomers = "select CustomerID from Customers where CustomerPrefix='" + Common.toString(AgentName) + "'";
                                                string result = DBUtils.executeSqlGetSingle(sqlCustomers);
                                                if (Common.toString(result).Trim() == "")
                                                {
                                                    //** get data from GlChartofAccounts (prefix, title, code , country, currency) and save in Customers
                                                    string getInfo = "Select Prefix,AccountName, AccountCode,CountryISONumericCode,CurrencyISOCode from GLChartOfAccounts where AccountCode ='" + AgentAccountCode + "'";
                                                    DataTable dtCustomer = DBUtils.GetDataTable(getInfo);
                                                    if (dtCustomer != null)
                                                    {
                                                        if (dtCustomer.Rows.Count > 0)
                                                        {
                                                            DataRow drCustomer = dtCustomer.Rows[0];
                                                            string createCustomer = "Insert into Customers(CustomerFirstName,AccountCode,CustomerCountryISOCode,CustomerCurrency,CustomerPrefix,CustomerCompany,Status) Values (@CustomerFirstName,@AccountCode,@CustomerCountryISOCode,@CustomerCurrency,@CustomerPrefix,@CustomerCompany,@Status)";
                                                            SqlCommand Insertcommand = null;
                                                            Insertcommand = con.CreateCommand();
                                                            Insertcommand.Parameters.AddWithValue("@CustomerPrefix", drCustomer["Prefix"]);
                                                            Insertcommand.Parameters.AddWithValue("@CustomerFirstName", drCustomer["AccountName"]);
                                                            Insertcommand.Parameters.AddWithValue("@AccountCode", drCustomer["AccountCode"]);
                                                            Insertcommand.Parameters.AddWithValue("@CustomerCountryISOCode", drCustomer["CountryISONumericCode"]);
                                                            Insertcommand.Parameters.AddWithValue("@CustomerCurrency", drCustomer["CurrencyISOCode"]);
                                                            Insertcommand.Parameters.AddWithValue("@CustomerCompany", drCustomer["AccountName"]);
                                                            Insertcommand.Parameters.AddWithValue("@Status", true);
                                                            Insertcommand.CommandText = createCustomer;
                                                            Insertcommand.ExecuteNonQuery();
                                                            Insertcommand.Parameters.Clear();
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Currency not not setup for agent<br>";
                                        }
                                    }
                                    else
                                    {
                                        IsRowValid = false;
                                        rowErrMessage += "Agent account head not setup.<br>";
                                    }
                                }
                                else
                                {
                                    //InsertCommand.Parameters.AddWithValue("@AgentName", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Agent name is not available in excel sheet<br>";
                                }
                            }

                            if (i == 2) // Customer_id
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    int CustomerId = Common.toInt(ws.Row(r).Cell(i).Value);
                                    InsertCommand.Parameters.AddWithValue("@CustomerId", CustomerId);
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@CustomerId", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "Customer id is not available in excel sheet<br>";
                                }
                            }

                            if (i == 3) // Customer_Name
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@CustomerName", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@CustomerName", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Customer name is not available in excel sheet<br>";
                                }
                            }

                            if (i == 4) // TRAN ID
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@TRANID", Common.toInt(ws.Row(r).Cell(i).Value));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@TRANID", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "Transaction Id is not available in excel sheet<br>";
                                }
                            }

                            if (i == 5) // Beneficiary_full_name
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Recipient", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Recipient", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Recipient is not available in excel sheet<br> ";
                                }
                            }


                            if (i == 6) // FatherName
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@FatherName", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@FatherName", "");
                                    //IsRowValid = false;
                                    //rowErrMessage += "FatherName is not available in excel sheet; ";
                                }
                            }
                            if (i == 7) // Phone
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Phone", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Phone", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Phone is not available in excel sheet<br> ";
                                }
                            }

                            if (i == 8) // Address
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Address", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Address", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Address is not available in excel sheet<br>";
                                }
                            }


                            if (i == 9) // City
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@City", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@City", "");
                                    IsRowValid = false;
                                    rowErrMessage += "City is not available in excel sheet<br>";
                                }
                            }

                            if (i == 10) // Code
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Code", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Code", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Code is not available in excel sheet<br> ";
                                }
                            }

                            if (i == 11) // PayOut_CCY
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    PayoutCurrency = Common.toString(ws.Row(r).Cell(i).Value);
                                    if (Common.toString(PayoutCurrency).Trim() != "")
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PayoutCurrency", PayoutCurrency);
                                        if (Common.isCurrencyAvailable(PayoutCurrency))
                                        {
                                            string accountcode = Common.getHoldAccountByCurrency(PayoutCurrency, "1");
                                            if (accountcode == "")
                                            {
                                                AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(PayoutCurrency, "1");
                                                if (!response.isAdded)
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += ws.Row(r).Cell(i).Value.ToString() + " Unpaid Account is not available and unable to create it. Please check currency settings.<br>";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += PayoutCurrency + " is not available or is inactive. Please check currency settings<br>";
                                        }
                                    }
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@PayoutCurrency", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Payout currency is not available in excel sheet<br>";
                                }
                            }

                            if (i == 12) // PayOut_Amount
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@FCAmount", Common.toDecimal(ws.Row(r).Cell(i).Value));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@FCAmount", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "F/C amount is not available in excel sheet<br>";
                                }
                            }
                            if (i == 13) // Agent_rate
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    decimal Rate = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                    //if (Common.toString(PayoutCurrency).Trim() != "" && Rate > 0)
                                    //{
                                    //    decimal MinRateLimit = 0m;
                                    //    decimal MiaxRateLimit = 0m;
                                    //    DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + PayoutCurrency + "'");
                                    //    if (objCurrency != null)
                                    //    {
                                    //        if (objCurrency.Rows.Count > 0)
                                    //        {
                                    //            DataRow drCurrency = objCurrency.Rows[0];
                                    //            MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                    //            MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                                    //        }
                                    //        if (Rate < MinRateLimit || Rate > MiaxRateLimit)
                                    //        {
                                    //            IsRowValid = false;
                                    //            rowErrMessage += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                                    //        }
                                    //    }
                                    //}
                                    InsertCommand.Parameters.AddWithValue("@Rate", Rate);
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Rate", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "Rate is not available in excel sheet<br>";
                                }
                            }
                            if (i == 14) // Payin_amount
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Payinamount", Common.toDecimal(ws.Row(r).Cell(i).Value));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Payinamount", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "Pay in amount is not available in excel sheet<br>";
                                }
                            }

                            if (i == 15) // PaymentNo
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@PaymentNo", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@PaymentNo", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Payment No is not available in excel sheet<br>";
                                }
                            }

                            if (i == 16) // TransactionDate
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@TransactionDate", Common.toDateTime(ws.Row(r).Cell(i).Value));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@TransactionDate", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Transaction date is not available in excel sheet<br>";
                                }
                            }
                            if (i == 17) // Admin_Charges
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    decimal AdminCharges = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                    InsertCommand.Parameters.AddWithValue("@AdminCharges", AdminCharges);
                                    ////if (AdminCharges > 0)
                                    ////{
                                    ////    if (AgentAccountCode != "" && AgentCurrencyCode != "")
                                    ////    {
                                    ////        string accountcode = Common.getHoldAccountByCurrency(AgentCurrencyCode, "2");
                                    ////        if (accountcode == "")
                                    ////        {
                                    ////            AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(AgentCurrencyCode, "2");
                                    ////            if (!response.isAdded)
                                    ////            {
                                    ////                IsRowValid = false;
                                    ////                rowErrMessage += AgentCurrencyCode + " Un-earned admin charges account is not available and unable to create it. Please check currency settings.<br>";
                                    ////            }
                                    ////        }
                                    ////    }
                                    ////}
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@AdminCharges", 0);
                                    //IsRowValid = false;
                                    //rowErrMessage += "Admin_Charges is not available in excel sheet;";
                                }
                            }
                            if (i == 18) // Agent_Charges
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    decimal AgentCharges = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                    InsertCommand.Parameters.AddWithValue("@AgentCharges", AgentCharges);
                                    ////if (AgentCharges > 0)
                                    ////{
                                    ////    if (AgentAccountCode != "" && AgentCurrencyCode != "")
                                    ////    {
                                    ////        string accountcode = Common.getHoldAccountByCurrency(AgentCurrencyCode, "3");
                                    ////        if (accountcode == "")
                                    ////        {
                                    ////            AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(AgentCurrencyCode, "3");
                                    ////            if (!response.isAdded)
                                    ////            {
                                    ////                IsRowValid = false;
                                    ////                rowErrMessage += AgentCurrencyCode + " Un-earned agent charges account is not available and unable to create it. Please check currency settings.<br>";
                                    ////            }
                                    ////        }
                                    ////    }
                                    ////}
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@AgentCharges", 0);
                                }
                            }

                            if (i == 19) // BenePaymentMethod
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@BenePaymentMethod", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@BenePaymentMethod", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Bene Payment Method is not available in excel sheet<br>";
                                }
                            }

                            if (i == 20) // SendingCountry
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@SendingCountry", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@SendingCountry", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Sending Country is not available in excel sheet<br>";
                                }
                            }

                            if (i == 21) // ReceivingCountry
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@ReceivingCountry", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@ReceivingCountry", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Receiving Country is not available in excel sheet<br>";
                                }
                            }
                            if (i == 22) // Status
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Status", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Status", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Status is not available in excel sheet.<br>";
                                }
                            }
                            if (i == 23) // Posting Date
                            {
                                if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        //is valid date
                                        //is within fin year
                                        if (Common.isDate(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                        {
                                            if (Common.ValidateFiscalYearDate(Common.toDateTime(ws.Row(r).Cell(i).Value)))
                                            {
                                                if (Common.ValidatePostingDate(Common.toString(ws.Row(r).Cell(i).Value)))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(ws.Row(r).Cell(i).Value));
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                    IsRowValid = false;
                                                    rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                                }
                                            }
                                            else
                                            {
                                                InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                IsRowValid = false;
                                                rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                                            }

                                        }
                                        else
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                            IsRowValid = false;
                                            rowErrMessage += "Posting date is not valid.<br>";
                                        }

                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Posting date is not available in excel sheet.<br>";
                                    }

                                }
                                else
                                {

                                    if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                    }
                                }
                            }
                        }

                        if (IsRowValid)
                        {
                            InsertCommand.Parameters.AddWithValue("@isRowValid", true);
                            InsertCommand.Parameters.AddWithValue("@cpUserRefNo", viewModel.cpUserRefNo);
                            InsertCommand.Parameters.AddWithValue("@RowErrMessage", rowErrMessage);
                        }
                        else
                        {
                            InsertCommand.Parameters.AddWithValue("@isRowValid", false);
                            InsertCommand.Parameters.AddWithValue("@RowErrMessage", rowErrMessage);
                            InsertCommand.Parameters.AddWithValue("@cpUserRefNo", viewModel.cpUserRefNo);
                        }
                        InsertCommand.Parameters.AddWithValue("@RowNumber", j);

                        InsertCommand.CommandText = sqlMyEntry;
                        InsertCommand.ExecuteNonQuery();
                        InsertCommand.Parameters.Clear();

                        rowErrMessage = "";
                        PayoutCurrency = "";
                        viewModel.isPosted = true;
                        j++;
                        //returnModel.isPosted = true;
                    }
                    con.Close();
                }
                else
                {
                    viewModel.ErrMessage = "Incomplete data in excel sheet.";
                }
            }
            catch (Exception ex)
            {
                viewModel.ErrMessage = "Unable to load data Please try Again";
            }
            return Json(viewModel, JsonRequestBehavior.AllowGet);
        }
        public ActionResult CheckList_Read([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM FileEntriesTemp /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                          Id,RowNumber,cpUserRefNo, TransactionDate,PostingDate, AgentName, Rate,PayoutCurrency,FCAmount,Payinamount,AdminCharges,AgentCharges,TRANID,CustomerId,CustomerName,FatherName,Recipient,Phone,Address,City,Code,PaymentNo,BenePaymentMethod,SendingCountry,ReceivingCountry,Status,isRowValid,RowErrMessage FROM FileEntriesTemp
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";
            //     const string selectQuery = @"SELECT
            //CheckListId, CheckTypeId, (select CheckType  from CheckTypes where CheckTypes.CheckTypeId = CheckList.CheckTypeId ) as CheckTypeName, CheckText, case when Status = 'true' then 'Active' else 'Inactive' end as StatusName, Status FROM CheckList";
            SqlBuilder builder = new SqlBuilder();
            var count = builder.AddTemplate(countQuery);

            int CurrentPage = 0; int CurrentPageSize = 1;
            if (request.PageSize > 0)
            {
                CurrentPageSize = request.PageSize;
            }
            if (request.PageSize > 0 && request.Page > 0)
            {
                CurrentPage = request.Page - 1;
            }
            var selector = builder.AddTemplate(selectQuery, new { PageIndex = CurrentPage, PageSize = CurrentPageSize });

            if (request.Filters != null && request.Filters.Any())
            {
                builder = Common.ApplyFilters(builder, request.Filters);
            }
            else
            {
                builder.Where("cpUserRefNo=''");
            }

            //
            //builder.Where("isHead='true'");

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("RowNumber");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);

            var rows = _dbcontext.Query<HoldFilePostingViewModel>(selector.RawSql, selector.Parameters);

            //  rows.Each(item => Console.WriteLine($"({ item.AccountCode}): { item.AccountName}
            //{                GetParentsString(items, item)}"));            
            //}
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult EditingInline_Update([DataSourceRequest] DataSourceRequest request, HoldFilePostingViewModel product)
        {
            if (product != null)
            {
                product = EditingCustom_Update(product);
            }

            return Json(new[] { product }.ToDataSourceResult(request, ModelState));
        }
        public HoldFilePostingViewModel EditingCustom_Update(HoldFilePostingViewModel pFileModel)
        {
            ApplicationUser Profile = new ApplicationUser();
            bool IsRowValid = true;
            string err_msgs = "";
            string UpdateSQL = "Update FileEntriesTemp Set ";
            UpdateSQL += "TransactionDate=@TransactionDate,PostingDate=@PostingDate,AgentName=@AgentName,AgentCharges=@AgentCharges,Rate=@Rate,PayoutCurrency=@PayoutCurrency,FCAmount=@FCAmount,Payinamount=@Payinamount, AdminCharges=@AdminCharges,TRANID=@TRANID,CustomerId=@CustomerId,CustomerName=@CustomerName,FatherName=@FatherName,Recipient=@Recipient,Phone=@Phone,Address=@Address,City=@City,Code=@Code,PaymentNo=@PaymentNo,BenePaymentMethod=@BenePaymentMethod,SendingCountry=@SendingCountry,ReceivingCountry=@ReceivingCountry,Status=@Status,isRowValid=@isRowValid,RowErrMessage=@RowErrMessage Where Id=@Id";
            using (SqlConnection con = Common.getConnection())
            {
                SqlCommand updateCommand = con.CreateCommand();
                try
                {
                    if (Common.toString(pFileModel.TransactionDate).Trim() != "")
                    {
                        updateCommand.Parameters.AddWithValue("@TransactionDate", pFileModel.TransactionDate);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Transaction date is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@TransactionDate", "");
                    }

                    if (Common.toString(pFileModel.PostingDate).Trim() != "")
                    {
                        if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                        {
                            if (Common.isDate(Common.toString(pFileModel.PostingDate)))
                            {
                                if (Common.ValidateFiscalYearDate(Common.toDateTime(pFileModel.PostingDate)))
                                {
                                    if (Common.ValidatePostingDate(Common.toString(pFileModel.PostingDate)))
                                    {
                                        updateCommand.Parameters.AddWithValue("@PostingDate", pFileModel.PostingDate);
                                    }
                                    else
                                    {
                                        updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                        IsRowValid = false;
                                        err_msgs += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                    }

                                }
                                else
                                {
                                    updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                    IsRowValid = false;
                                    err_msgs += "Posting date is not between financial year start date and financial year end date.<br>";
                                }

                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Posting date is not valid.<br>";
                            }

                        }
                        else
                        {
                            if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", SysPrefs.PostingDate);
                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                            }
                        }

                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Posting date is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@PostingDate", "");
                    }

                    if (!string.IsNullOrEmpty(pFileModel.AgentName))
                    {
                        updateCommand.Parameters.AddWithValue("@AgentName", pFileModel.AgentName);
                        string accountcode = Common.getGLAccountCodeByPrefix(pFileModel.AgentName);
                        if (accountcode.Trim() != "")
                        {
                            string sqlCustomers = "select CustomerID from Customers where CustomerPrefix='" + Common.toString(pFileModel.AgentName) + "'";
                            string result = DBUtils.executeSqlGetSingle(sqlCustomers);
                            if (result.ToString().Trim() == "")
                            {
                                IsRowValid = false;
                                err_msgs += "Agent/customer does not exists in the system.<br>";
                            }
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Invalid prefix. Agent prefix not available in chart of account<br>";
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Agent Name is not available in excel sheet<br>";
                        updateCommand.Parameters.AddWithValue("@AgentName", "");
                    }
                    // Agent_rate
                    if (pFileModel.Rate > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Rate", pFileModel.Rate);
                        //if (!string.IsNullOrEmpty(pFileModel.PayoutCurrency) )
                        //{
                        //    decimal MinRateLimit = 0m;
                        //    decimal MiaxRateLimit = 0m;
                        //    DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + pFileModel.PayoutCurrency + "'.<br>");
                        //    if (objCurrency != null)
                        //    {
                        //        if (objCurrency.Rows.Count > 0)
                        //        {
                        //            DataRow drCurrency = objCurrency.Rows[0];
                        //            MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                        //            MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                        //        }
                        //        if (pFileModel.Rate < MinRateLimit || pFileModel.Rate > MiaxRateLimit)
                        //        {
                        //            IsRowValid = false;
                        //            err_msgs += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                        //        }
                        //    }
                        //}
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Rate is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Rate", "");
                    }
                    // PayOut_CCY
                    if (!string.IsNullOrEmpty(pFileModel.PayoutCurrency))
                    {
                        updateCommand.Parameters.AddWithValue("@PayoutCurrency", pFileModel.PayoutCurrency);
                        string PayoutCurrency = pFileModel.PayoutCurrency;
                        if (Common.toString(PayoutCurrency).Trim() != "")
                        {
                            if (Common.isCurrencyAvailable(PayoutCurrency))
                            {
                                string accountcode = Common.getHoldAccountByCurrency(PayoutCurrency, "1");
                                if (accountcode == "")
                                {
                                    AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(PayoutCurrency, "1");
                                    if (!response.isAdded)
                                    {
                                        IsRowValid = false;
                                        err_msgs += PayoutCurrency + " Unpaid Account is not available and unable to create it. Please check currency settings.<br>";
                                    }
                                }
                            }
                            else
                            {
                                IsRowValid = false;
                                err_msgs += PayoutCurrency + " is not available or is inactive. Please check currency settings.<br>";
                            }
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Payout currency is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@PayoutCurrency", "");
                    }

                    // Amount
                    if (pFileModel.FCAmount > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@FCAmount", pFileModel.FCAmount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "F/C amount is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@FCAmount", "0");
                    }
                    // Payin_amount
                    if (pFileModel.Payinamount >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Payinamount", pFileModel.Payinamount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Payin amount is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Payinamount", "0");
                    }
                    /// Admin Charges
                    if (pFileModel.AdminCharges >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@AdminCharges", pFileModel.AdminCharges);
                        ////if (pFileModel.AdminCharges > 0 && Common.toString(pFileModel.PayoutCurrency) != "")
                        ////{
                        ////    string accountcode = Common.getHoldAccountByCurrency(pFileModel.PayoutCurrency, "2");
                        ////    if (accountcode == "")
                        ////    {
                        ////        AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(pFileModel.PayoutCurrency, "2");
                        ////        if (!response.isAdded)
                        ////        {
                        ////            IsRowValid = false;
                        ////            err_msgs += pFileModel.PayoutCurrency + " Un-earned admin charges account is not available and unable to create it. Please check currency settings.<br>";
                        ////        }
                        ////    }
                        ////}
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@AdminCharges", "0");
                    }

                    ////     Agent Charges 
                    if (pFileModel.AgentCharges >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@AgentCharges", pFileModel.AgentCharges);
                        ////if (pFileModel.AgentCharges > 0 && Common.toString(pFileModel.PayoutCurrency) != "")
                        ////{
                        ////    string accountcode = Common.getHoldAccountByCurrency(pFileModel.PayoutCurrency, "3");
                        ////    if (accountcode == "")
                        ////    {
                        ////        AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(pFileModel.PayoutCurrency, "3");
                        ////        if (!response.isAdded)
                        ////        {
                        ////            IsRowValid = false;
                        ////            err_msgs += pFileModel.PayoutCurrency + " Un-earned agent charges account is not available and unable to create it. Please check currency settings.<br>";
                        ////        }
                        ////    }
                        ////}
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@AgentCharges", "0");
                    }
                    /// Phone
                    if (!string.IsNullOrEmpty(pFileModel.Phone))
                    {
                        updateCommand.Parameters.AddWithValue("@Phone", pFileModel.Phone);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Phone is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Phone", "");
                    }
                    /// FatherName
                    if (!string.IsNullOrEmpty(pFileModel.FatherName))
                    {
                        updateCommand.Parameters.AddWithValue("@FatherName", pFileModel.FatherName);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "FatherName is not available in excel sheet;";
                        updateCommand.Parameters.AddWithValue("@FatherName", "");
                    }
                    /// Address
                    if (!string.IsNullOrEmpty(pFileModel.Address))
                    {
                        updateCommand.Parameters.AddWithValue("@Address", pFileModel.Address);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Address is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Address", "");
                    }
                    /// City
                    if (!string.IsNullOrEmpty(pFileModel.City))
                    {
                        updateCommand.Parameters.AddWithValue("@City", pFileModel.City);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "City is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@City", "");
                    }
                    /// Code
                    if (!string.IsNullOrEmpty(pFileModel.Code))
                    {
                        updateCommand.Parameters.AddWithValue("@Code", pFileModel.Code);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Code is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Code", "");
                    }
                    /// PaymentNo
                    if (!string.IsNullOrEmpty(pFileModel.PaymentNo))
                    {
                        updateCommand.Parameters.AddWithValue("@PaymentNo", pFileModel.PaymentNo);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "PaymentNo is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@PaymentNo", "");
                    }
                    /// BenePaymentMethod
                    if (!string.IsNullOrEmpty(pFileModel.BenePaymentMethod))
                    {
                        updateCommand.Parameters.AddWithValue("@BenePaymentMethod", pFileModel.BenePaymentMethod);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Bene Payment Method is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@BenePaymentMethod", "");
                    }
                    /// SendingCountry
                    if (!string.IsNullOrEmpty(pFileModel.SendingCountry))
                    {
                        updateCommand.Parameters.AddWithValue("@SendingCountry", pFileModel.SendingCountry);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Sending Country is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@SendingCountry", "");
                    }
                    /// ReceivingCountry
                    if (!string.IsNullOrEmpty(pFileModel.ReceivingCountry))
                    {
                        updateCommand.Parameters.AddWithValue("@ReceivingCountry", pFileModel.ReceivingCountry);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Receiving Country is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@ReceivingCountry", "");
                    }
                    /// Tr_No
                    if (pFileModel.TRANID > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@TRANID", pFileModel.TRANID);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Transaction id is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@TRANID", "");
                    }
                    /// Customer_id
                    if (pFileModel.CustomerId > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerId", pFileModel.CustomerId);
                        string sqlCustomers = "select CustomerID from Customers where CustomerId=" + pFileModel.CustomerId + "";
                        string result = DBUtils.executeSqlGetSingle(sqlCustomers);
                        if (result.ToString().Trim() == "")
                        {
                            IsRowValid = false;
                            err_msgs += "Agent/customer does not exists in the system.<br>";
                            //Agent/customer does not exists in the system
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Customer id is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@CustomerId", "");
                    }
                    /// Customer_full_name
                    if (!string.IsNullOrEmpty(pFileModel.CustomerName))
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerName", pFileModel.CustomerName);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Customer name is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@CustomerName", "");
                    }
                    /// Beneficiary_full_name
                    if (!string.IsNullOrEmpty(pFileModel.Recipient))
                    {
                        updateCommand.Parameters.AddWithValue("@Recipient", pFileModel.Recipient);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Recipient is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Recipient", "");
                    }

                    if (!string.IsNullOrEmpty(pFileModel.Status))
                    {
                        updateCommand.Parameters.AddWithValue("@Status", pFileModel.Status);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Status is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Status", "");
                    }

                    if (IsRowValid)
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", true);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                        updateCommand.Parameters.AddWithValue("@RowErrMessage", err_msgs);
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", false);
                        updateCommand.Parameters.AddWithValue("@RowErrMessage", err_msgs);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                    }
                    updateCommand.CommandText = UpdateSQL;
                    updateCommand.ExecuteNonQuery();
                    updateCommand.Parameters.Clear();

                }
                catch (Exception ex)
                {

                }
                finally
                {
                    con.Close();
                }
            }
            pFileModel.isRowValid = IsRowValid;
            pFileModel.RowErrMessage = err_msgs;
            return pFileModel;
        }
        public ActionResult PostHoldData(HoldFilePostingViewModel pPostingModel)
        {
            ErrorMessages objErrorMessage = null;
            ////string myConnectionString = BaseModel.getConnString();////
            string myConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
            ApplicationUser profile = ApplicationUser.GetUserProfile();
            DataSet DS_lstGLEntriesTemp = DBUtils.GetDataSet("select * from FileEntriesTemp where cpUserRefNo='" + pPostingModel.cpUserRefNo + "'");
            DataTable lstGLEntriesTemp = DS_lstGLEntriesTemp.Tables[0];
            if (lstGLEntriesTemp != null)
            {

                if (lstGLEntriesTemp.Rows.Count > 0)
                {
                    bool isMainValid = true;
                    foreach (DataRow objGLEntriesTemp in lstGLEntriesTemp.Rows)
                    {
                        if (Common.toBool(objGLEntriesTemp["isRowValid"]) == false)
                        {
                            isMainValid = false;
                        }
                    }
                    if (isMainValid)
                    {
                        objErrorMessage = PostingUtils.SaveHoldTransactionBridge("22HOLDENTRY", lstGLEntriesTemp, Profile.Id.ToString());
                        string postingErr = "";
                        if (objErrorMessage.IsFailed)
                        {
                            for (int i = 0; i < objErrorMessage.ErrorMessage.Count; i++)
                            {
                                postingErr = postingErr + "<br>" + objErrorMessage.ErrorMessage[i].Message;
                            }
                            pPostingModel.ErrMessage = Common.GetAlertMessage(1, postingErr);
                        }
                        if (objErrorMessage.ErrorMessage.Count > 0)
                        {
                            pPostingModel.ErrMessage = objErrorMessage.ErrorMessage[0].Message.ToString();
                        }
                        else
                        {
                            pPostingModel.ErrMessage = Common.GetAlertMessage(0, "Hold Entry from Excel is successfully added <br> <br> " + objErrorMessage.ErrorMessage[0].Message + "  <br> <br> <a href ='/Inquiries/GLInquiry'>Click here</a>");
                        }
                    }
                    else
                    {
                        pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Please check your data and try again");
                    }
                }
            }
            else
            {
                pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Data not found");
            }

            return View("~/Views/ImportExcel/Message.cshtml", pPostingModel);
        }

        #endregion

        #region Import Paid File
        public ActionResult PaidFile()
        {
            DBUtils.ExecuteSQL("Delete from FileEntriesTemp where (cpUserRefNo = '') or (cpUserRefNo is null) or (cpUserRefNo='0')");
            ImportPaidFileViewModel postingViewModel = new ImportPaidFileViewModel();
            Guid guid = Guid.NewGuid();
            postingViewModel.cpUserRefNo = guid.ToString();
            if (string.IsNullOrEmpty(SysPrefs.DefaultExchangeVariancesAccount))
            {
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Unable to find default exchange variances account. Please update preferences.");
            }
            if (string.IsNullOrEmpty(SysPrefs.DefaultCurrency))
            {
                postingViewModel.ErrMessage += Common.GetAlertMessage(1, "Unable to find default curreny. Please update preferences.");
            }
            if (string.IsNullOrEmpty(SysPrefs.PostingDate))
            {
                postingViewModel.ErrMessage += Common.GetAlertMessage(1, "Unable to find posting date. Please update preferences.");
            }
            return View("~/Views/ImportExcel/PaidFile.cshtml", postingViewModel);
        }

        [HttpPost]
        public ActionResult PaidFileUploaded(ImportPaidFileViewModel pImportPaidFileViewModel, HttpPostedFileBase file)
        {
            try
            {
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                IEnumerable<Sheets> pcheckListViewModel = new List<Sheets>();
                var fileName = ""; var path = "";
                pImportPaidFileViewModel.isFileImported = false;
                if (file != null && file.ContentLength > 0)
                {
                    var supportedTypes = new[] { "xlsx" };
                    var fileExt = Path.GetExtension(file.FileName).Substring(1);
                    if (!supportedTypes.Contains(fileExt))
                    {
                        pImportPaidFileViewModel.ErrMessage = Common.GetAlertMessage(1, "File extension is invalid -  only xlsx files are allowed.");
                        pImportPaidFileViewModel.isError = true;
                    }
                    else
                    {

                        fileName = Path.GetFileName(file.FileName);
                        ////fileName = Common.getRandomDigit(6) + "-" + fileName;
                        fileName = "Paid-" + DateTime.Now.ToString("MMddyyyyHHmmssffff") + ".xlsx";
                        path = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/UpTemp/"), fileName);
                        file.SaveAs(path);
                        pImportPaidFileViewModel.isFileImported = true;

                        pImportPaidFileViewModel.ExcelFilePath = path;

                        if (path != null && path != "")
                        {
                            int Count = 0;
                            var workbook = new XLWorkbook(path);
                            foreach (var worksheet in workbook.Worksheets)
                            {
                                var sheetname = worksheet.ToString().Split('!');
                                var sheetname1 = sheetname[0].ToString();
                                sheetname1 = RemoveSpecialCharacters(sheetname1);
                                Sheets pSheets = new Sheets();
                                pSheets.Name = Common.toString(sheetname1);
                                pSheets.Name = pSheets.Name.Replace("'", "");
                                pSheets.Value = Count.ToString();
                                pcheckListViewModel = (new[] { pSheets }).Concat(pcheckListViewModel);
                                Count++;
                            }

                            pImportPaidFileViewModel.sheets = pcheckListViewModel.ToList();

                        }
                        else
                        {
                            pImportPaidFileViewModel.ErrMessage = "file not uploaded. please try again.";
                            pImportPaidFileViewModel.isError = true;
                        }
                    }
                }
                else
                {
                    pImportPaidFileViewModel.ErrMessage = "file not uploaded. please try again.";
                    pImportPaidFileViewModel.isError = true;
                }
            }
            catch (Exception ex)
            {
                pImportPaidFileViewModel.ErrMessage = "An error occured: " + ex.ToString();
                pImportPaidFileViewModel.isError = true;
            }
            //return ddlSheets;
            return View("~/Views/ImportExcel/PaidFile.cshtml", pImportPaidFileViewModel);
        }

        public ActionResult ShowPaidExcelfiledata(ImportPaidFileViewModel viewModel)
        {
            string sqlMyEntry = "";
            try
            {
                string savepath = viewModel.ExcelFilePath;
                var workbook = new XLWorkbook(savepath);
                IXLWorksheet ws;
                string PayoutCurrency = "";
                bool IsRowValid = true;
                string rowErrMessage = "";

                string sheetname = viewModel.Sheet;
                workbook.TryGetWorksheet(Common.toString(sheetname), out ws);
                var datarange = ws.RangeUsed();
                int TotalRowsinExcelSheet = datarange.RowCount();
                int TotalColinExcelSheet = datarange.ColumnCount();
                int TotalFileColumns = 26;
                if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                {
                    TotalFileColumns = 27;
                }
                if (TotalColinExcelSheet == TotalFileColumns)
                {
                    int row_number = 0;
                    if (viewModel.HasHeader)
                    {
                        row_number = 2;
                    }
                    else
                    {
                        row_number = 1;
                    }
                    int j = 1;
                    SqlCommand InsertCommand = null;
                    SqlConnection con = Common.getConnection();
                    for (int r = row_number; r <= TotalRowsinExcelSheet; r++)
                    {
                        sqlMyEntry = "Insert into FileEntriesTemp (RowNumber,TransactionDate,AgentName,Rate,PayoutCurrency,FCAmount,Payinamount, AdminCharges,AgentCharges,TRANID,CustomerId,CustomerName,FatherName,Recipient,Phone,Address,City,Code,BuyerRate,Buyer,BuyerRateSC,BuyerRateDC,PaymentNo,BenePaymentMethod,SendingCountry,ReceivingCountry,Status, isRowValid, RowErrMessage,cpUserRefNo,PostingDate) Values (@RowNumber,@TransactionDate,@AgentName,@Rate,@PayoutCurrency,@FCAmount,@Payinamount, @AdminCharges,@AgentCharges,@TRANID,@CustomerId,@CustomerName,@FatherName,@Recipient,@Phone,@Address,@City,@Code,@BuyerRate,@Buyer,@BuyerRateSC,@BuyerRateDC,@PaymentNo,@BenePaymentMethod,@SendingCountry,@ReceivingCountry,@Status,@isRowValid,@RowErrMessage,@cpUserRefNo,@PostingDate)";

                        string AgentName = "";
                        string AgentAccountCode = "";
                        string AgentCurrencyCode = "";
                        string TransactionId = "";
                        string BuyerName = "";
                        InsertCommand = con.CreateCommand();
                        if (ws.Row(r).Cell(26).Value.ToString() != "Canceled" && ws.Row(r).Cell(26).Value.ToString() != "Refunded")
                        {
                            if (Common.GetSysSettings("ImportFileHasPostingDate") == "0")
                            {
                                if (Common.ValidateFiscalYearDate(Common.toDateTime(SysPrefs.PostingDate)))
                                {
                                    if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                    }
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                                }
                            }
                            for (int i = 1; i <= TotalColinExcelSheet; i++)
                            {
                                if (i == 1) // Agent Name
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        AgentName = Common.toString(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@AgentName", AgentName);
                                        AgentAccountCode = Common.getGLAccountCodeByPrefix(AgentName);
                                        if (AgentAccountCode.Trim() != "")
                                        {
                                            AgentCurrencyCode = Common.getGLAccountCurrencyByCode(AgentAccountCode);
                                            if (AgentCurrencyCode != "")
                                            {
                                                if (!Common.isCurrencyAvailable(AgentCurrencyCode))
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += AgentCurrencyCode + " is not available or is inactive. Please check currency settings<br>";
                                                    AgentCurrencyCode = "";
                                                }
                                            }
                                            else
                                            {
                                                IsRowValid = false;
                                                rowErrMessage += "Currency not not setup for agent<br>";
                                            }

                                            string sqlCustomers = "select CustomerID from Customers where CustomerPrefix='" + Common.toString(AgentName) + "'";
                                            string result = DBUtils.executeSqlGetSingle(sqlCustomers);
                                            if (result.ToString().Trim() == "")
                                            {
                                                //** get data from GlChartofAccounts (prefix, title, code , country, currency) and save in Customers
                                                //IsRowValid = false;
                                                //rowErrMessage += "Agent/customer does not exists in the system.<br>";
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Invalid prefix. Agent prefix not available in chart of account<br>";
                                        }
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@AgentName", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Agent name is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 2) // Customer_id
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        int CustomerId = Common.toInt(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@CustomerId", CustomerId);
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@CustomerId", 0);
                                        //IsRowValid = false;
                                        //rowErrMessage += "Customer id is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 3) // Customer_Name
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@CustomerName", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@CustomerName", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Customer name is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 4) // TRAN ID
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        TransactionId = Common.toString(ws.Row(r).Cell(i).Value).Trim();
                                        string sqlTranx = "select TransactionID from CustomerTransactions where TransactionID='" + TransactionId + "'";
                                        string result = DBUtils.executeSqlGetSingle(sqlTranx);
                                        if (Common.toString(result).Trim() == "")
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Can not find this transaction (" + TransactionId + ") in unpaid/hold transaction list.<br>";
                                        }
                                        InsertCommand.Parameters.AddWithValue("@TRANID", TransactionId);
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@TRANID", 0);
                                        IsRowValid = false;
                                        rowErrMessage += "Transaction Id is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 5) // Beneficiary_full_name
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Recipient", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Recipient", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Recipient is not available in excel sheet<br> ";
                                    }
                                }


                                if (i == 6) // FatherName
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@FatherName", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@FatherName", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "FatherName is not available in excel sheet; ";
                                    }
                                }
                                if (i == 7) // Phone
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Phone", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Phone", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Phone is not available in excel sheet<br> ";
                                    }
                                }

                                if (i == 8) // Address
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Address", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Address", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Address is not available in excel sheet<br>";
                                    }
                                }


                                if (i == 9) // City
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@City", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@City", "");
                                        ///IsRowValid = false;
                                        //rowErrMessage += "City is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 10) // Code
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Code", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Code", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Code is not available in excel sheet<br> ";
                                    }
                                }

                                if (i == 11) // PayOut_CCY
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        PayoutCurrency = Common.toString(ws.Row(r).Cell(i).Value);
                                        if (Common.toString(PayoutCurrency).Trim() != "")
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PayoutCurrency", PayoutCurrency);
                                            if (Common.isCurrencyAvailable(PayoutCurrency))
                                            {
                                                string accountcode = Common.getHoldAccountByCurrency(PayoutCurrency, "1");
                                                if (accountcode == "")
                                                {
                                                    AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(PayoutCurrency, "1");
                                                    if (!response.isAdded)
                                                    {
                                                        IsRowValid = false;
                                                        rowErrMessage += ws.Row(r).Cell(i).Value.ToString() + " Unpaid Account is not available and unable to create it. Please check currency settings.<br>";
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                IsRowValid = false;
                                                rowErrMessage += PayoutCurrency + " is not available or is inactive. Please check currency settings<br>";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PayoutCurrency", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Payout currency is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 12) // PayOut_Amount
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@FCAmount", Common.toDecimal(ws.Row(r).Cell(i).Value));
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@FCAmount", 0);
                                        IsRowValid = false;
                                        rowErrMessage += "F/C amount is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 13) // Agent_rate
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        decimal Rate = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                        //if (Common.toString(PayoutCurrency).Trim() != "" && Rate > 0)
                                        //{
                                        //    decimal MinRateLimit = 0m;
                                        //    decimal MiaxRateLimit = 0m;
                                        //    DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + PayoutCurrency + "'");
                                        //    if (objCurrency != null)
                                        //    {
                                        //        if (objCurrency.Rows.Count > 0)
                                        //        {
                                        //            DataRow drCurrency = objCurrency.Rows[0];
                                        //            MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                        //            MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                                        //        }
                                        //        if (Rate < MinRateLimit || Rate > MiaxRateLimit)
                                        //        {
                                        //            IsRowValid = false;
                                        //            rowErrMessage += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                                        //        }
                                        //    }
                                        //}
                                        InsertCommand.Parameters.AddWithValue("@Rate", Rate);
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Rate", 0);
                                        IsRowValid = false;
                                        rowErrMessage += "Rate is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 14) // Payin_amount
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Payinamount", Common.toDecimal(ws.Row(r).Cell(i).Value));
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Payinamount", 0);
                                        IsRowValid = false;
                                        rowErrMessage += "Pay in amount is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 15) // PaymentNo
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        string sqlPaymentNo = "select PaymentNo from CustomerTransactions where TransactionID='" + TransactionId + "' and PaymentNo='" + Common.toString(ws.Row(r).Cell(i).Value).Trim() + "'";
                                        string result = DBUtils.executeSqlGetSingle(sqlPaymentNo);
                                        if (Common.toString(result).Trim() == "")
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Can not find this payment # (" + Common.toString(ws.Row(r).Cell(i).Value) + ") in unpaid/hold transaction list.<br>";
                                        }
                                        InsertCommand.Parameters.AddWithValue("@PaymentNo", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PaymentNo", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Payment No is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 16) // TransactionDate
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@TransactionDate", Common.toDateTime(ws.Row(r).Cell(i).Value));
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@TransactionDate", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Transaction date is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 17) // Admin_Charges
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        decimal AdminCharges = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@AdminCharges", AdminCharges);
                                        ////if (AdminCharges > 0)
                                        ////{
                                        ////    if (AgentAccountCode != "" && AgentCurrencyCode != "")
                                        ////    {
                                        ////        string accountcode = Common.getHoldAccountByCurrency(AgentCurrencyCode, "2");
                                        ////        if (accountcode == "")
                                        ////        {
                                        ////            AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(AgentCurrencyCode, "2");
                                        ////            if (!response.isAdded)
                                        ////            {
                                        ////                IsRowValid = false;
                                        ////                rowErrMessage += AgentCurrencyCode + " Un-earned admin charges account is not available and unable to create it. Please check currency settings.<br>";
                                        ////            }
                                        ////        }
                                        ////    }
                                        ////}
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@AdminCharges", 0);
                                        //IsRowValid = false;
                                        //rowErrMessage += "Admin_Charges is not available in excel sheet;";
                                    }
                                }
                                if (i == 18) // Agent_Charges
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        decimal AgentCharges = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@AgentCharges", AgentCharges);
                                        ////if (AgentCharges > 0)
                                        ////{
                                        ////    if (AgentAccountCode != "" && AgentCurrencyCode != "")
                                        ////    {
                                        ////        string accountcode = Common.getHoldAccountByCurrency(AgentCurrencyCode, "3");
                                        ////        if (accountcode == "")
                                        ////        {
                                        ////            AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(AgentCurrencyCode, "3");
                                        ////            if (!response.isAdded)
                                        ////            {
                                        ////                IsRowValid = false;
                                        ////                rowErrMessage += AgentCurrencyCode + " Un-earned agent charges account is not available and unable to create it. Please check currency settings.<br>";
                                        ////            }
                                        ////        }
                                        ////    }
                                        ////}
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@AgentCharges", 0);
                                    }
                                }

                                if (i == 19) // Buyer Rate
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        decimal BuyerRate = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                        if (Common.toString(PayoutCurrency).Trim() != "" && BuyerRate > 0)
                                        {
                                            decimal MinRateLimit = 0m;
                                            decimal MiaxRateLimit = 0m;
                                            DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + PayoutCurrency + "'");
                                            if (objCurrency != null)
                                            {
                                                if (objCurrency.Rows.Count > 0)
                                                {
                                                    DataRow drCurrency = objCurrency.Rows[0];
                                                    MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                                    MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                                                }
                                                if (BuyerRate < MinRateLimit || BuyerRate > MiaxRateLimit)
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Buyer rate should be > 0<br>";
                                        }
                                        InsertCommand.Parameters.AddWithValue("@BuyerRate", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRate", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Buyer rate is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 20) // Buyer
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        BuyerName = Common.toString(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@Buyer", BuyerName);
                                        AgentAccountCode = Common.getGLAccountCodeByPrefix(BuyerName);
                                        if (AgentAccountCode.Trim() != "")
                                        {
                                            AgentCurrencyCode = Common.getGLAccountCurrencyByCode(AgentAccountCode);
                                            if (AgentCurrencyCode != "")
                                            {
                                                if (!Common.isCurrencyAvailable(AgentCurrencyCode))
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += AgentCurrencyCode + " is not available or is inactive. Please check currency settings<br>";
                                                    AgentCurrencyCode = "";
                                                }
                                            }
                                            else
                                            {
                                                IsRowValid = false;
                                                rowErrMessage += "Currency not not setup for buyer<br>";
                                            }

                                            string sqlVendors = "select VendorID from Vendors where VendorPrefix ='" + Common.toString(BuyerName) + "'";
                                            string result = DBUtils.executeSqlGetSingle(sqlVendors);
                                            if (result.ToString().Trim() == "")
                                            {
                                                //** get data from GlChartofAccounts (prefix, title, code , country, currency) and save in Vendors
                                                string getInfo = "Select Prefix,AccountName, AccountCode,CountryISONumericCode,CurrencyISOCode from GLChartOfAccounts where AccountCode ='" + AgentAccountCode + "'";
                                                DataTable dtVendor = DBUtils.GetDataTable(getInfo);
                                                if (dtVendor != null)
                                                {
                                                    if (dtVendor.Rows.Count > 0)
                                                    {
                                                        DataRow drVendor = dtVendor.Rows[0];
                                                        try
                                                        {
                                                            string createVendor = "Insert into Vendors(VendorPrefix,VendorFirstName,AccountCode,VendorCountryISOCode,VendorCurrency,VendorCompany,Status) Values (@VendorPrefix,@VendorFirstName,@AccountCode,@VendorCountryISOCode,@VendorCurrency,@VendorCompany,@Status)";
                                                            SqlCommand Insertcommand = null;
                                                            Insertcommand = con.CreateCommand();
                                                            Insertcommand.Parameters.AddWithValue("@VendorPrefix", Common.toString(drVendor["Prefix"]));
                                                            Insertcommand.Parameters.AddWithValue("@VendorFirstName", Common.toString(drVendor["AccountName"]));
                                                            Insertcommand.Parameters.AddWithValue("@AccountCode", Common.toString(drVendor["AccountCode"]));
                                                            Insertcommand.Parameters.AddWithValue("@VendorCountryISOCode", Common.toString(drVendor["CountryISONumericCode"]));
                                                            Insertcommand.Parameters.AddWithValue("@VendorCurrency", Common.toString(drVendor["CurrencyISOCode"]));
                                                            Insertcommand.Parameters.AddWithValue("@VendorCompany", Common.toString(drVendor["AccountName"]));
                                                            Insertcommand.Parameters.AddWithValue("@Status", true);
                                                            Insertcommand.CommandText = createVendor;
                                                            Insertcommand.ExecuteNonQuery();
                                                            Insertcommand.Parameters.Clear();
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            string error = Common.toString(drVendor["Prefix"]) + "<br/>";
                                                            error += Common.toString(drVendor["AccountName"]) + "<br/>";
                                                            error += Common.toString(drVendor["AccountCode"]) + "<br/>";
                                                            error += Common.toString(drVendor["CountryISONumericCode"]) + "<br/>";
                                                            error += Common.toString(drVendor["CurrencyISOCode"]) + "<br/>";


                                                            Response.Write(error + "  " + ex.ToString());
                                                            Response.End();
                                                        }
                                                    }
                                                }

                                                //IsRowValid = false;
                                                //rowErrMessage += "Agent/customer does not exists in the system.<br>";                                        }
                                            }
                                            else
                                            {
                                                //IsRowValid = false;
                                                //rowErrMessage += "Invalid prefix. Buyer prefix not available in chart of account<br>";
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Buyer is not available in excel sheet<br>";
                                        }
                                    }
                                }
                                if (i == 21) // Buyer Rate SC
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRateSC", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRateSC", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Buyer Rate SC is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 22) // Buyer Rate DC
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRateDC", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRateDC", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Buyer Rate DC is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 23) // BenePaymentMethod
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BenePaymentMethod", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BenePaymentMethod", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Bene Payment Method is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 24) // SendingCountry
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@SendingCountry", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@SendingCountry", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Sending Country is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 25) // ReceivingCountry
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@ReceivingCountry", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@ReceivingCountry", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Receiving Country is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 26) // Status
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Status", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Status", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Status is not available in excel sheet.<br>";
                                    }
                                }
                                if (i == 27) // Posting Date
                                {
                                    if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                                    {
                                        if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                        {
                                            //is valid date
                                            //is within fin year
                                            var pDate = ws.Row(r).Cell(i).Value;
                                            if (Common.isDate(Common.toString(pDate)))
                                            {
                                                if (Common.ValidateFiscalYearDate(Common.toDateTime(ws.Row(r).Cell(i).Value)))
                                                {
                                                    if (Common.ValidatePostingDate(Common.toString(pDate)))
                                                    {
                                                        InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(pDate));
                                                    }
                                                    else
                                                    {
                                                        InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                        IsRowValid = false;
                                                        rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                                    }
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                    IsRowValid = false;
                                                    rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                                                }

                                            }
                                            else
                                            {
                                                InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                IsRowValid = false;
                                                rowErrMessage += "Posting date is not valid.<br>";
                                            }

                                        }
                                        else
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                            IsRowValid = false;
                                            rowErrMessage += "Posting date is not available in excel sheet.<br>";
                                        }

                                    }
                                    else
                                    {
                                        if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                                        }
                                        else
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                            IsRowValid = false;
                                            rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                        }
                                    }
                                }
                            }

                            if (IsRowValid)
                            {
                                InsertCommand.Parameters.AddWithValue("@isRowValid", true);
                                InsertCommand.Parameters.AddWithValue("@cpUserRefNo", viewModel.cpUserRefNo);
                                InsertCommand.Parameters.AddWithValue("@RowErrMessage", rowErrMessage);
                            }
                            else
                            {
                                InsertCommand.Parameters.AddWithValue("@isRowValid", false);
                                InsertCommand.Parameters.AddWithValue("@RowErrMessage", rowErrMessage);
                                InsertCommand.Parameters.AddWithValue("@cpUserRefNo", viewModel.cpUserRefNo);
                            }

                            InsertCommand.Parameters.AddWithValue("@RowNumber", j);

                            InsertCommand.CommandText = sqlMyEntry;
                            InsertCommand.ExecuteNonQuery();
                            InsertCommand.Parameters.Clear();
                        }
                        rowErrMessage = "";
                        PayoutCurrency = "";
                        viewModel.isPosted = true;
                        j++;
                        //returnModel.isPosted = true;
                    }
                    con.Close();
                }
                else
                {
                    viewModel.ErrMessage = "Incomplete data in excel sheet.";
                }
            }
            catch (Exception ex)
            {
                viewModel.ErrMessage = "Unable to load data. Please try again. :" + ex.ToString();
            }
            return Json(viewModel, JsonRequestBehavior.AllowGet);
        }

        public ActionResult CheckList_ReadPaid([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM FileEntriesTemp /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                          Id,RowNumber,cpUserRefNo, TransactionDate, AgentName,PostingDate, Rate,PayoutCurrency,FCAmount,Payinamount,AdminCharges,AgentCharges,TRANID,CustomerId,CustomerName,FatherName,Recipient,Phone,Address,City,Code,BuyerRate,Buyer,BuyerRateSC,BuyerRateDC,PaymentNo,BenePaymentMethod,SendingCountry,ReceivingCountry,Status,isRowValid,RowErrMessage FROM FileEntriesTemp
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";
            //     const string selectQuery = @"SELECT
            //CheckListId, CheckTypeId, (select CheckType  from CheckTypes where CheckTypes.CheckTypeId = CheckList.CheckTypeId ) as CheckTypeName, CheckText, case when Status = 'true' then 'Active' else 'Inactive' end as StatusName, Status FROM CheckList";
            SqlBuilder builder = new SqlBuilder();
            var count = builder.AddTemplate(countQuery);

            int CurrentPage = 0; int CurrentPageSize = 1;
            if (request.PageSize > 0)
            {
                CurrentPageSize = request.PageSize;
            }
            if (request.PageSize > 0 && request.Page > 0)
            {
                CurrentPage = request.Page - 1;
            }
            var selector = builder.AddTemplate(selectQuery, new { PageIndex = CurrentPage, PageSize = CurrentPageSize });

            if (request.Filters != null && request.Filters.Any())
            {
                builder = Common.ApplyFilters(builder, request.Filters);
            }
            else
            {
                builder.Where("cpUserRefNo=''");
            }

            //cpUserRefNo
            //builder.Where("isHead='true'");

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("RowNumber");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);

            var rows = _dbcontext.Query<ImportPaidFileViewModel>(selector.RawSql, selector.Parameters);

            //  rows.Each(item => Console.WriteLine($"({ item.AccountCode}): { item.AccountName}
            //{                GetParentsString(items, item)}"));            
            //}
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult EditingInline_UpdatePaid([DataSourceRequest] DataSourceRequest request, ImportPaidFileViewModel product)
        {
            if (product != null)
            {
                product = EditingCustom_UpdatePaid(product);
            }

            return Json(new[] { product }.ToDataSourceResult(request, ModelState));
        }
        public ImportPaidFileViewModel EditingCustom_UpdatePaid(ImportPaidFileViewModel pFileModel)
        {
            ApplicationUser Profile = new ApplicationUser();
            bool IsRowValid = true;
            string err_msgs = "";
            string UpdateSQL = "Update FileEntriesTemp Set ";
            UpdateSQL += "TransactionDate=@TransactionDate,PostingDate=@PostingDate,AgentName=@AgentName,AgentCharges=@AgentCharges,Rate=@Rate,PayoutCurrency=@PayoutCurrency,FCAmount=@FCAmount,Payinamount=@Payinamount, AdminCharges=@AdminCharges,TRANID=@TRANID,CustomerId=@CustomerId,CustomerName=@CustomerName,FatherName=@FatherName,Recipient=@Recipient,Phone=@Phone,Address=@Address,City=@City,Code=@Code,BuyerRate=@BuyerRate,Buyer=@Buyer,BuyerRateSC=@BuyerRateSC,BuyerRateDC=@BuyerRateDC,PaymentNo=@PaymentNo,BenePaymentMethod=@BenePaymentMethod,SendingCountry=@SendingCountry,ReceivingCountry=@ReceivingCountry,Status=@Status,isRowValid=@isRowValid,RowErrMessage=@RowErrMessage Where Id=@Id";
            using (SqlConnection con = Common.getConnection())
            {
                SqlCommand updateCommand = con.CreateCommand();
                try
                {
                    if (Common.toString(pFileModel.TransactionDate).Trim() != "")
                    {
                        updateCommand.Parameters.AddWithValue("@TransactionDate", pFileModel.TransactionDate);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Transaction date is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@TransactionDate", "");
                    }

                    if (Common.toString(pFileModel.PostingDate).Trim() != "")
                    {
                        if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                        {
                            if (Common.isDate(Common.toString(pFileModel.PostingDate)))
                            {
                                if (Common.ValidateFiscalYearDate(Common.toDateTime(pFileModel.PostingDate)))
                                {
                                    if (Common.ValidatePostingDate(Common.toString(pFileModel.PostingDate)))
                                    {
                                        updateCommand.Parameters.AddWithValue("@PostingDate", pFileModel.PostingDate);
                                    }
                                    else
                                    {
                                        updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                        IsRowValid = false;
                                        err_msgs += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                    }
                                }
                                else
                                {
                                    updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                    IsRowValid = false;
                                    err_msgs += "Posting date is not between financial year start date and financial year end date.<br>";
                                }

                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Posting date is not valid.<br>";
                            }

                        }
                        else
                        {
                            if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", SysPrefs.PostingDate);
                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                            }
                        }

                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Posting date is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@PostingDate", "");
                    }

                    if (Common.toString(pFileModel.Buyer).Trim() != "")
                    {
                        string BuyerName = pFileModel.Buyer;
                        string AgentAccountCode = Common.getGLAccountCodeByPrefix(BuyerName);
                        if (AgentAccountCode.Trim() != "")
                        {
                            string AgentCurrencyCode = Common.getGLAccountCurrencyByCode(AgentAccountCode);
                            if (AgentCurrencyCode != "")
                            {
                                if (!Common.isCurrencyAvailable(AgentCurrencyCode))
                                {
                                    IsRowValid = false;
                                    err_msgs += AgentCurrencyCode + " is not available or is inactive. Please check currency settings<br>";
                                    AgentCurrencyCode = "";
                                }
                            }
                            else
                            {
                                IsRowValid = false;
                                err_msgs += "Currency not not setup for buyer<br>";
                            }

                            ////string sqlBuyers = "select VendorID from Vendors where VendorPrefix='" + Common.toString(BuyerName) + "'";
                            ////string result = DBUtils.executeSqlGetSingle(sqlBuyers);
                            ////if (result.ToString().Trim() == "")
                            ////{
                            ////    IsRowValid = false;
                            ////    err_msgs += "Buyer does not exists in the system.<br>";
                            ////}
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Invalid prefix. Buyer prefix not available in chart of account<br>";
                        }
                        updateCommand.Parameters.AddWithValue("@Buyer", pFileModel.Buyer);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Buyer is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@Buyer", "");
                    }

                    if (Common.toString(pFileModel.BuyerRate).Trim() != "")
                    {
                        decimal BuyerRate = pFileModel.BuyerRate;
                        if (Common.toString(pFileModel.PayoutCurrency).Trim() != "" && BuyerRate > 0)
                        {
                            decimal MinRateLimit = 0m;
                            decimal MiaxRateLimit = 0m;
                            DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + pFileModel.PayoutCurrency + "'");
                            if (objCurrency != null)
                            {
                                if (objCurrency.Rows.Count > 0)
                                {
                                    DataRow drCurrency = objCurrency.Rows[0];
                                    MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                    MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                                }
                                if (BuyerRate < MinRateLimit || BuyerRate > MiaxRateLimit)
                                {
                                    IsRowValid = false;
                                    err_msgs += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                                }
                            }
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Buyer rate should be > 0<br>";
                        }
                        updateCommand.Parameters.AddWithValue("@BuyerRate", pFileModel.BuyerRate);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Buyer rate is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@BuyerRate", "");
                    }

                    if (Common.toString(pFileModel.BuyerRateSC).Trim() != "")
                    {
                        updateCommand.Parameters.AddWithValue("@BuyerRateSC", pFileModel.BuyerRateSC);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Buyer Rate SC is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@BuyerRateSC", "");
                    }

                    if (Common.toString(pFileModel.BuyerRateDC).Trim() != "")
                    {
                        updateCommand.Parameters.AddWithValue("@BuyerRateDC", pFileModel.BuyerRateDC);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Buyer Rate DC is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@BuyerRateDC", "");
                    }

                    if (!string.IsNullOrEmpty(pFileModel.AgentName))
                    {
                        updateCommand.Parameters.AddWithValue("@AgentName", pFileModel.AgentName);
                        string accountcode = Common.getGLAccountCodeByPrefix(pFileModel.AgentName);
                        if (accountcode.Trim() != "")
                        {
                            //string sqlCustomers = "select CustomerID from Customers where CustomerPrefix='" + Common.toString(pFileModel.AgentName) + "'";
                            //string result = DBUtils.executeSqlGetSingle(sqlCustomers);
                            //if (result.ToString().Trim() == "")
                            //{
                            //    IsRowValid = false;
                            //    err_msgs += "Agent/customer does not exists in the system.<br>";
                            //}
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Invalid prefix. Agent prefix not available in chart of account<br>";
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Agent Name is not available in editable grid<br>";
                        updateCommand.Parameters.AddWithValue("@AgentName", "");
                    }
                    // Agent_rate
                    if (pFileModel.Rate > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Rate", pFileModel.Rate);
                        //if (!string.IsNullOrEmpty(pFileModel.PayoutCurrency) )
                        //{
                        //    decimal MinRateLimit = 0m;
                        //    decimal MiaxRateLimit = 0m;
                        //    DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + pFileModel.PayoutCurrency + "'.<br>");
                        //    if (objCurrency != null)
                        //    {
                        //        if (objCurrency.Rows.Count > 0)
                        //        {
                        //            DataRow drCurrency = objCurrency.Rows[0];
                        //            MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                        //            MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                        //        }
                        //        if (pFileModel.Rate < MinRateLimit || pFileModel.Rate > MiaxRateLimit)
                        //        {
                        //            IsRowValid = false;
                        //            err_msgs += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                        //        }
                        //    }
                        //}
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Rate is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@Rate", "");
                    }
                    // PayOut_CCY
                    if (!string.IsNullOrEmpty(pFileModel.PayoutCurrency))
                    {
                        updateCommand.Parameters.AddWithValue("@PayoutCurrency", pFileModel.PayoutCurrency);
                        string PayoutCurrency = pFileModel.PayoutCurrency;
                        if (Common.toString(PayoutCurrency).Trim() != "")
                        {
                            if (Common.isCurrencyAvailable(PayoutCurrency))
                            {
                                string accountcode = Common.getHoldAccountByCurrency(PayoutCurrency, "1");
                                if (accountcode == "")
                                {
                                    AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(PayoutCurrency, "1");
                                    if (!response.isAdded)
                                    {
                                        IsRowValid = false;
                                        err_msgs += PayoutCurrency + " Unpaid Account is not available and unable to create it. Please check currency settings.<br>";
                                    }
                                }
                            }
                            else
                            {
                                IsRowValid = false;
                                err_msgs += PayoutCurrency + " is not available or is inactive. Please check currency settings.<br>";
                            }
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Payout currency is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@PayoutCurrency", "");
                    }

                    // Amount
                    if (pFileModel.FCAmount > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@FCAmount", pFileModel.FCAmount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "F/C amount is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@FCAmount", "0");
                    }
                    // Payin_amount
                    if (pFileModel.Payinamount >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Payinamount", pFileModel.Payinamount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Payin amount is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@Payinamount", "0");
                    }
                    /// Admin Charges
                    if (pFileModel.AdminCharges >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@AdminCharges", pFileModel.AdminCharges);
                        ////if (pFileModel.AdminCharges > 0 && Common.toString(pFileModel.PayoutCurrency) != "")
                        ////{
                        ////    string accountcode = Common.getHoldAccountByCurrency(pFileModel.PayoutCurrency, "2");
                        ////    if (accountcode == "")
                        ////    {
                        ////        AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(pFileModel.PayoutCurrency, "2");
                        ////        if (!response.isAdded)
                        ////        {
                        ////            IsRowValid = false;
                        ////            err_msgs += pFileModel.PayoutCurrency + " Un-earned admin charges account is not available and unable to create it. Please check currency settings.<br>";
                        ////        }
                        ////    }
                        ////}
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@AdminCharges", "0");
                    }

                    ////     Agent Charges 
                    if (pFileModel.AgentCharges >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@AgentCharges", pFileModel.AgentCharges);
                        ////if (pFileModel.AgentCharges > 0 && Common.toString(pFileModel.PayoutCurrency) != "")
                        ////{
                        ////    string accountcode = Common.getHoldAccountByCurrency(pFileModel.PayoutCurrency, "3");
                        ////    if (accountcode == "")
                        ////    {
                        ////        AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(pFileModel.PayoutCurrency, "3");
                        ////        if (!response.isAdded)
                        ////        {
                        ////            IsRowValid = false;
                        ////            err_msgs += pFileModel.PayoutCurrency + " Un-earned agent charges account is not available and unable to create it. Please check currency settings.<br>";
                        ////        }
                        ////    }
                        ////}
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@AgentCharges", "0");
                    }
                    /// Phone
                    if (!string.IsNullOrEmpty(pFileModel.Phone))
                    {
                        updateCommand.Parameters.AddWithValue("@Phone", pFileModel.Phone);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Phone is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Phone", "");
                    }
                    /// FatherName
                    if (!string.IsNullOrEmpty(pFileModel.FatherName))
                    {
                        updateCommand.Parameters.AddWithValue("@FatherName", pFileModel.FatherName);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "FatherName is not available in excel sheet;";
                        updateCommand.Parameters.AddWithValue("@FatherName", "");
                    }
                    /// Address
                    if (!string.IsNullOrEmpty(pFileModel.Address))
                    {
                        updateCommand.Parameters.AddWithValue("@Address", pFileModel.Address);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Address is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Address", "");
                    }
                    /// City
                    if (!string.IsNullOrEmpty(pFileModel.City))
                    {
                        updateCommand.Parameters.AddWithValue("@City", pFileModel.City);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "City is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@City", "");
                    }
                    /// Code
                    if (!string.IsNullOrEmpty(pFileModel.Code))
                    {
                        updateCommand.Parameters.AddWithValue("@Code", pFileModel.Code);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Code is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Code", "");
                    }
                    /// PaymentNo
                    if (!string.IsNullOrEmpty(pFileModel.PaymentNo))
                    {
                        updateCommand.Parameters.AddWithValue("@PaymentNo", pFileModel.PaymentNo);
                        if (!string.IsNullOrEmpty(Common.toString(pFileModel.PaymentNo).Trim()))
                        {
                            string sqlPaymentNo = "select PaymentNo from CustomerTransactions where PaymentNo='" + Common.toString(pFileModel.PaymentNo).Trim() + "'";
                            string result = DBUtils.executeSqlGetSingle(sqlPaymentNo);
                            if (Common.toString(result).Trim() == "")
                            {
                                IsRowValid = false;
                                err_msgs += "Can not find this payment # (" + Common.toString(pFileModel.PaymentNo) + ") in unpaid/hold transaction list.<br>";
                            }
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Payment No is not available in editable grid<br>";
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "PaymentNo is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@PaymentNo", "");
                    }
                    /// BenePaymentMethod
                    if (!string.IsNullOrEmpty(pFileModel.BenePaymentMethod))
                    {
                        updateCommand.Parameters.AddWithValue("@BenePaymentMethod", pFileModel.BenePaymentMethod);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Bene Payment Method is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@BenePaymentMethod", "");
                    }
                    /// SendingCountry
                    if (!string.IsNullOrEmpty(pFileModel.SendingCountry))
                    {
                        updateCommand.Parameters.AddWithValue("@SendingCountry", pFileModel.SendingCountry);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Sending Country is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@SendingCountry", "");
                    }
                    /// ReceivingCountry
                    if (!string.IsNullOrEmpty(pFileModel.ReceivingCountry))
                    {
                        updateCommand.Parameters.AddWithValue("@ReceivingCountry", pFileModel.ReceivingCountry);
                    }
                    else
                    {
                        //   IsRowValid = false;
                        //err_msgs += "Receiving Country is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@ReceivingCountry", "");
                    }
                    /// Tr_No
                    if (pFileModel.TRANID > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@TRANID", pFileModel.TRANID);
                        if (!string.IsNullOrEmpty(Common.toString(pFileModel.TRANID).Trim()))
                        {
                            string sqlTranx = "select TransactionID from CustomerTransactions where TransactionID='" + pFileModel.TRANID + "'";
                            string result = DBUtils.executeSqlGetSingle(sqlTranx);
                            if (Common.toString(result).Trim() == "")
                            {
                                IsRowValid = false;
                                err_msgs += "Can not find this transaction (" + pFileModel.TRANID + ") in unpaid/hold transaction list.<br>";
                            }
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Transaction Id is not available in editable grid<br>";
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Transaction id is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@TRANID", "");
                    }
                    /// Customer_id
                    if (pFileModel.CustomerId > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerId", pFileModel.CustomerId);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Customer id is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@CustomerId", "");
                    }
                    /// Customer_full_name
                    if (!string.IsNullOrEmpty(pFileModel.CustomerName))
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerName", pFileModel.CustomerName);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Customer name is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@CustomerName", "");
                    }
                    /// Beneficiary_full_name
                    if (!string.IsNullOrEmpty(pFileModel.Recipient))
                    {
                        updateCommand.Parameters.AddWithValue("@Recipient", pFileModel.Recipient);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Recipient is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Recipient", "");
                    }

                    if (!string.IsNullOrEmpty(pFileModel.Status))
                    {
                        updateCommand.Parameters.AddWithValue("@Status", pFileModel.Status);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Status is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Status", "");
                    }

                    if (IsRowValid)
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", true);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                        updateCommand.Parameters.AddWithValue("@RowErrMessage", err_msgs);
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", false);
                        updateCommand.Parameters.AddWithValue("@RowErrMessage", err_msgs);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                    }
                    updateCommand.CommandText = UpdateSQL;
                    updateCommand.ExecuteNonQuery();
                    updateCommand.Parameters.Clear();

                }
                catch (Exception ex)
                {

                }
                finally
                {
                    con.Close();
                }
            }

            pFileModel.isRowValid = IsRowValid;
            pFileModel.RowErrMessage = err_msgs;
            return pFileModel;
        }

        public ActionResult PostPaidData(ImportPaidFileViewModel pPostingModel)
        {
            ErrorMessages objErrorMessage = null;
            ////  string myConnectionString = BaseModel.getConnString();//// System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
            ApplicationUser profile = ApplicationUser.GetUserProfile();
            DataSet DS_lstGLEntriesTemp = DBUtils.GetDataSet("select * from FileEntriesTemp where cpUserRefNo='" + pPostingModel.cpUserRefNo + "'");
            DataTable lstGLEntriesTemp = DS_lstGLEntriesTemp.Tables[0];
            if (lstGLEntriesTemp != null)
            {

                if (lstGLEntriesTemp.Rows.Count > 0)
                {
                    bool isMainValid = true;
                    foreach (DataRow objGLEntriesTemp in lstGLEntriesTemp.Rows)
                    {
                        if (Common.toBool(objGLEntriesTemp["isRowValid"]) == false)
                        {
                            isMainValid = false;
                        }
                    }
                    if (isMainValid)
                    {
                        objErrorMessage = PostingUtils.SavePaidTransactionBridge("22PAIDENTRY", lstGLEntriesTemp, Profile.Id.ToString());
                        string postingErr = "";
                        if (objErrorMessage.IsFailed)
                        {
                            for (int i = 0; i < objErrorMessage.ErrorMessage.Count; i++)
                            {
                                postingErr = postingErr + "<br>" + objErrorMessage.ErrorMessage[i].Message;
                            }
                            pPostingModel.ErrMessage = Common.GetAlertMessage(1, postingErr);
                        }
                        if (objErrorMessage.ErrorMessage.Count > 0)
                        {
                            pPostingModel.ErrMessage = objErrorMessage.ErrorMessage[0].Message.ToString();
                        }
                        else
                        {
                            pPostingModel.ErrMessage = Common.GetAlertMessage(0, "Paid transaction from excel is successfully added <br> <br> " + objErrorMessage.ErrorMessage[0].Message + "  <br> <br> <a href ='/Inquiries/GLInquiry'>Click here</a>");
                        }
                    }
                    else
                    {
                        pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Please check your data and try again");
                    }
                }
            }
            else
            {
                pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Data not found");
            }
            HoldFilePostingViewModel pModel = new HoldFilePostingViewModel();
            pModel.ErrMessage = pPostingModel.ErrMessage;

            return View("~/Views/ImportExcel/Message.cshtml", pModel);
        }
        #endregion

        #region Import Branch Paid File
        public ActionResult BranchPaidFile()
        {
            DBUtils.ExecuteSQL("Delete from FileEntriesTemp where (cpUserRefNo = '') or (cpUserRefNo is null) or (cpUserRefNo='0')");
            ImportPaidFileViewModel postingViewModel = new ImportPaidFileViewModel();
            Guid guid = Guid.NewGuid();
            postingViewModel.cpUserRefNo = guid.ToString();
            if (string.IsNullOrEmpty(SysPrefs.DefaultExchangeVariancesAccount))
            {
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Unable to find default exchange variances account. Please update preferences.");
            }
            if (string.IsNullOrEmpty(SysPrefs.DefaultCurrency))
            {
                postingViewModel.ErrMessage += Common.GetAlertMessage(1, "Unable to find default curreny. Please update preferences.");
            }
            if (string.IsNullOrEmpty(SysPrefs.PostingDate))
            {
                postingViewModel.ErrMessage += Common.GetAlertMessage(1, "Unable to find posting date. Please update preferences.");
            }
            return View("~/Views/ImportExcel/BranchPaidFile.cshtml", postingViewModel);
        }

        [HttpPost]
        public ActionResult BranchPaidFileUploaded(ImportPaidFileViewModel pImportPaidFileViewModel, HttpPostedFileBase file)
        {
            try
            {
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                IEnumerable<Sheets> pcheckListViewModel = new List<Sheets>();
                var fileName = ""; var path = "";
                pImportPaidFileViewModel.isFileImported = false;
                if (file != null && file.ContentLength > 0)
                {
                    var supportedTypes = new[] { "xlsx" };
                    var fileExt = Path.GetExtension(file.FileName).Substring(1);
                    if (!supportedTypes.Contains(fileExt))
                    {
                        pImportPaidFileViewModel.ErrMessage = Common.GetAlertMessage(1, "File extension is invalid -  only xlsx files are allowed.");
                        pImportPaidFileViewModel.isError = true;
                    }
                    else
                    {
                        fileName = Path.GetFileName(file.FileName);
                        //// fileName = Common.getRandomDigit(6) + "-" + fileName;

                        //// path = Path.Combine(Server.MapPath("/Uploads/" + SysPrefs.SubmissionFolder + "/"), fileName);
                        fileName = "BranchPaid-" + DateTime.Now.ToString("MMddyyyyHHmmssffff") + ".xlsx";
                        path = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/UpTemp/"), fileName);
                        file.SaveAs(path);
                        pImportPaidFileViewModel.isFileImported = true;
                        pImportPaidFileViewModel.ExcelFilePath = path;

                        if (path != null && path != "")
                        {
                            int Count = 0;
                            var workbook = new XLWorkbook(path);
                            foreach (var worksheet in workbook.Worksheets)
                            {
                                var sheetname = worksheet.ToString().Split('!');
                                var sheetname1 = sheetname[0].ToString();
                                sheetname1 = RemoveSpecialCharacters(sheetname1);

                                Sheets pSheets = new Sheets();
                                pSheets.Name = Common.toString(sheetname1);
                                pSheets.Name = pSheets.Name.Replace("'", "");
                                pSheets.Value = Count.ToString();
                                pcheckListViewModel = (new[] { pSheets }).Concat(pcheckListViewModel);
                                Count++;
                            }

                            pImportPaidFileViewModel.sheets = pcheckListViewModel.ToList();

                        }
                        else
                        {
                            pImportPaidFileViewModel.ErrMessage = "file not uploaded. please try again.";
                            pImportPaidFileViewModel.isError = true;
                        }
                    }
                }
                else
                {
                    pImportPaidFileViewModel.ErrMessage = "file not uploaded. please try again.";
                    pImportPaidFileViewModel.isError = true;
                }
            }
            catch (Exception ex)
            {
                pImportPaidFileViewModel.ErrMessage = "An error occured: " + ex.ToString();
                pImportPaidFileViewModel.isError = true;
            }
            //return ddlSheets;
            return View("~/Views/ImportExcel/BranchPaidFile.cshtml", pImportPaidFileViewModel);
        }

        public ActionResult ShowBranchPaidExcelFileData(ImportPaidFileViewModel viewModel)
        {
            string sqlMyEntry = "";
            try
            {
                string savepath = viewModel.ExcelFilePath;
                var workbook = new XLWorkbook(savepath);
                IXLWorksheet ws;
                string PayoutCurrency = "";
                bool IsRowValid = true;
                string rowErrMessage = "";

                string sheetname = viewModel.Sheet;
                workbook.TryGetWorksheet(Common.toString(sheetname), out ws);
                var datarange = ws.RangeUsed();
                int TotalRowsinExcelSheet = datarange.RowCount();
                int TotalColinExcelSheet = datarange.ColumnCount();
                int TotalFileColumns = 26;
                if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                {
                    TotalFileColumns = 27;
                }
                if (TotalColinExcelSheet == TotalFileColumns)
                {
                    int row_number = 0;
                    if (viewModel.HasHeader)
                    {
                        row_number = 2;
                    }
                    else
                    {
                        row_number = 1;
                    }
                    int j = 1;
                    SqlCommand InsertCommand = null;
                    SqlConnection con = Common.getConnection();
                    for (int r = row_number; r <= TotalRowsinExcelSheet; r++)
                    {
                        sqlMyEntry = "Insert into FileEntriesTemp (RowNumber,TransactionDate,AgentName,Rate,PayoutCurrency,FCAmount,Payinamount, AdminCharges,AgentCharges,TRANID,CustomerId,CustomerName,FatherName,Recipient,Phone,Address,City,Code,BuyerRate,Buyer,BuyerRateSC,BuyerRateDC,PaymentNo,BenePaymentMethod,SendingCountry,ReceivingCountry,Status, isRowValid, RowErrMessage,cpUserRefNo,PostingDate) Values (@RowNumber,@TransactionDate,@AgentName,@Rate,@PayoutCurrency,@FCAmount,@Payinamount, @AdminCharges,@AgentCharges,@TRANID,@CustomerId,@CustomerName,@FatherName,@Recipient,@Phone,@Address,@City,@Code,@BuyerRate,@Buyer,@BuyerRateSC,@BuyerRateDC,@PaymentNo,@BenePaymentMethod,@SendingCountry,@ReceivingCountry,@Status,@isRowValid,@RowErrMessage,@cpUserRefNo,@PostingDate)";
                        ////  SqlCommand InsertCommand = null;
                        ////  SqlConnection con = Common.getConnection();
                        string AgentName = "";
                        string AgentAccountCode = "";
                        string AgentCurrencyCode = "";
                        string TransactionId = "";
                        string BuyerName = "";
                        InsertCommand = con.CreateCommand();
                        if (ws.Row(r).Cell(26).Value.ToString() != "Canceled" && ws.Row(r).Cell(26).Value.ToString() != "Refunded")
                        {
                            if (Common.GetSysSettings("ImportFileHasPostingDate") == "0")
                            {
                                if (Common.ValidateFiscalYearDate(Common.toDateTime(SysPrefs.PostingDate)))
                                {
                                    InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                                }
                            }
                            for (int i = 1; i <= TotalColinExcelSheet; i++)
                            {
                                if (i == 1) // Agent Name
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        AgentName = Common.toString(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@AgentName", AgentName);
                                        AgentAccountCode = Common.getGLAccountCodeByPrefix(AgentName);
                                        if (AgentAccountCode.Trim() != "")
                                        {
                                            AgentCurrencyCode = Common.getGLAccountCurrencyByCode(AgentAccountCode);
                                            if (AgentCurrencyCode != "")
                                            {
                                                if (!Common.isCurrencyAvailable(AgentCurrencyCode))
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += AgentCurrencyCode + " is not available or is inactive. Please check currency settings<br>";
                                                    AgentCurrencyCode = "";
                                                }
                                            }
                                            else
                                            {
                                                IsRowValid = false;
                                                rowErrMessage += "Currency not not setup for agent<br>";
                                            }

                                            string sqlCustomers = "select CustomerID from Customers where CustomerPrefix='" + Common.toString(AgentName) + "'";
                                            string result = DBUtils.executeSqlGetSingle(sqlCustomers);
                                            if (result.ToString().Trim() == "")
                                            {
                                                //** get data from GlChartofAccounts (prefix, title, code , country, currency) and save in Customers
                                                //IsRowValid = false;
                                                //rowErrMessage += "Agent/customer does not exists in the system.<br>";

                                                string getInfo = "Select Prefix,AccountName, AccountCode,CountryISONumericCode,CurrencyISOCode from GLChartOfAccounts where AccountCode ='" + AgentAccountCode + "'";
                                                DataTable dtCustomer = DBUtils.GetDataTable(getInfo);
                                                if (dtCustomer != null)
                                                {
                                                    if (dtCustomer.Rows.Count > 0)
                                                    {
                                                        DataRow drCustomer = dtCustomer.Rows[0];
                                                        string createCustomer = "Insert into Customers(CustomerFirstName,AccountCode,CustomerCountryISOCode,CustomerCurrency,CustomerPrefix,CustomerCompany,Status) Values (@CustomerFirstName,@AccountCode,@CustomerCountryISOCode,@CustomerCurrency,@CustomerPrefix,@CustomerCompany,@Status)";
                                                        SqlCommand Insertcommand = null;
                                                        Insertcommand = con.CreateCommand();
                                                        Insertcommand.Parameters.AddWithValue("@CustomerPrefix", drCustomer["Prefix"]);
                                                        Insertcommand.Parameters.AddWithValue("@CustomerFirstName", drCustomer["AccountName"]);
                                                        Insertcommand.Parameters.AddWithValue("@AccountCode", drCustomer["AccountCode"]);
                                                        Insertcommand.Parameters.AddWithValue("@CustomerCountryISOCode", drCustomer["CountryISONumericCode"]);
                                                        Insertcommand.Parameters.AddWithValue("@CustomerCurrency", drCustomer["CurrencyISOCode"]);
                                                        Insertcommand.Parameters.AddWithValue("@CustomerCompany", drCustomer["AccountName"]);
                                                        Insertcommand.Parameters.AddWithValue("@Status", true);
                                                        Insertcommand.CommandText = createCustomer;
                                                        Insertcommand.ExecuteNonQuery();
                                                        Insertcommand.Parameters.Clear();
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Invalid prefix. Agent prefix not available in chart of account<br>";
                                        }
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@AgentName", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Agent name is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 2) // Customer_id
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        int CustomerId = Common.toInt(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@CustomerId", CustomerId);
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@CustomerId", 0);
                                        //IsRowValid = false;
                                        //rowErrMessage += "Customer id is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 3) // Customer_Name
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@CustomerName", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@CustomerName", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Customer name is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 4) // TRAN ID
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        TransactionId = Common.toString(ws.Row(r).Cell(i).Value).Trim();
                                        //string sqlTranx = "select TransactionID from CustomerTransactions where TransactionID='" + TransactionId + "'";
                                        //string result = DBUtils.executeSqlGetSingle(sqlTranx);
                                        //if (Common.toString(result).Trim() == "")
                                        //{
                                        //    IsRowValid = false;
                                        //    rowErrMessage += "Can not find this transaction (" + TransactionId + ") in unpaid/hold transaction list.<br>";
                                        //}
                                        InsertCommand.Parameters.AddWithValue("@TRANID", TransactionId);
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@TRANID", 0);
                                        IsRowValid = false;
                                        rowErrMessage += "Transaction Id is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 5) // Beneficiary_full_name
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Recipient", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Recipient", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Recipient is not available in excel sheet<br> ";
                                    }
                                }


                                if (i == 6) // FatherName
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@FatherName", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@FatherName", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "FatherName is not available in excel sheet; ";
                                    }
                                }
                                if (i == 7) // Phone
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Phone", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Phone", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Phone is not available in excel sheet<br> ";
                                    }
                                }

                                if (i == 8) // Address
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Address", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Address", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Address is not available in excel sheet<br>";
                                    }
                                }


                                if (i == 9) // City
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@City", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@City", "");
                                        ///IsRowValid = false;
                                        //rowErrMessage += "City is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 10) // Code
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Code", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Code", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Code is not available in excel sheet<br> ";
                                    }
                                }

                                if (i == 11) // PayOut_CCY
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        PayoutCurrency = Common.toString(ws.Row(r).Cell(i).Value);
                                        if (Common.toString(PayoutCurrency).Trim() != "")
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PayoutCurrency", PayoutCurrency);
                                            if (Common.isCurrencyAvailable(PayoutCurrency))
                                            {
                                                string accountcode = Common.getHoldAccountByCurrency(PayoutCurrency, "1");
                                                if (accountcode == "")
                                                {
                                                    AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(PayoutCurrency, "1");
                                                    if (!response.isAdded)
                                                    {
                                                        IsRowValid = false;
                                                        rowErrMessage += ws.Row(r).Cell(i).Value.ToString() + " Unpaid Account is not available and unable to create it. Please check currency settings.<br>";
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                IsRowValid = false;
                                                rowErrMessage += PayoutCurrency + " is not available or is inactive. Please check currency settings<br>";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PayoutCurrency", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Payout currency is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 12) // PayOut_Amount
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@FCAmount", Common.toDecimal(ws.Row(r).Cell(i).Value));
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@FCAmount", 0);
                                        IsRowValid = false;
                                        rowErrMessage += "F/C amount is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 13) // Agent_rate
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        decimal Rate = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                        //if (Common.toString(PayoutCurrency).Trim() != "" && Rate > 0)
                                        //{
                                        //    decimal MinRateLimit = 0m;
                                        //    decimal MiaxRateLimit = 0m;
                                        //    DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + PayoutCurrency + "'");
                                        //    if (objCurrency != null)
                                        //    {
                                        //        if (objCurrency.Rows.Count > 0)
                                        //        {
                                        //            DataRow drCurrency = objCurrency.Rows[0];
                                        //            MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                        //            MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                                        //        }
                                        //        if (Rate < MinRateLimit || Rate > MiaxRateLimit)
                                        //        {
                                        //            IsRowValid = false;
                                        //            rowErrMessage += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                                        //        }
                                        //    }
                                        //}
                                        InsertCommand.Parameters.AddWithValue("@Rate", Rate);
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Rate", 0);
                                        IsRowValid = false;
                                        rowErrMessage += "Rate is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 14) // Payin_amount
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Payinamount", Common.toDecimal(ws.Row(r).Cell(i).Value));

                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Payinamount", 0);
                                        IsRowValid = false;
                                        rowErrMessage += "Pay in amount is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 15) // PaymentNo
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        ////string sqlPaymentNo = "select PaymentNo from CustomerTransactions where TransactionID='" + TransactionId + "' and PaymentNo='" + Common.toString(ws.Row(r).Cell(i).Value).Trim() + "'";
                                        ////string result = DBUtils.executeSqlGetSingle(sqlPaymentNo);
                                        ////if (Common.toString(result).Trim() == "")
                                        ////{
                                        ////    IsRowValid = false;
                                        ////    rowErrMessage += "Can not find this payment # (" + Common.toString(ws.Row(r).Cell(i).Value) + ") in unpaid/hold transaction list.<br>";
                                        ////}
                                        InsertCommand.Parameters.AddWithValue("@PaymentNo", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PaymentNo", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Payment No is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 16) // TransactionDate
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@TransactionDate", Common.toDateTime(ws.Row(r).Cell(i).Value));
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@TransactionDate", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Transaction date is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 17) // Admin_Charges
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        decimal AdminCharges = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@AdminCharges", AdminCharges);
                                        ////if (AdminCharges > 0)
                                        ////{
                                        ////    if (AgentAccountCode != "" && AgentCurrencyCode != "")
                                        ////    {
                                        ////        string accountcode = Common.getHoldAccountByCurrency(AgentCurrencyCode, "2");
                                        ////        if (accountcode == "")
                                        ////        {
                                        ////            AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(AgentCurrencyCode, "2");
                                        ////            if (!response.isAdded)
                                        ////            {
                                        ////                IsRowValid = false;
                                        ////                rowErrMessage += AgentCurrencyCode + " Un-earned admin charges account is not available and unable to create it. Please check currency settings.<br>";
                                        ////            }
                                        ////        }
                                        ////    }
                                        ////}
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@AdminCharges", 0);
                                        //IsRowValid = false;
                                        //rowErrMessage += "Admin_Charges is not available in excel sheet;";
                                    }
                                }
                                if (i == 18) // Agent_Charges
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        decimal AgentCharges = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@AgentCharges", AgentCharges);
                                        ////if (AgentCharges > 0)
                                        ////{
                                        ////    if (AgentAccountCode != "" && AgentCurrencyCode != "")
                                        ////    {
                                        ////        string accountcode = Common.getHoldAccountByCurrency(AgentCurrencyCode, "3");
                                        ////        if (accountcode == "")
                                        ////        {
                                        ////            AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(AgentCurrencyCode, "3");
                                        ////            if (!response.isAdded)
                                        ////            {
                                        ////                IsRowValid = false;
                                        ////                rowErrMessage += AgentCurrencyCode + " Un-earned agent charges account is not available and unable to create it. Please check currency settings.<br>";
                                        ////            }
                                        ////        }
                                        ////    }
                                        ////}
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@AgentCharges", 0);
                                    }
                                }

                                if (i == 19) // Buyer Rate
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        decimal BuyerRate = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                        if (Common.toString(PayoutCurrency).Trim() != "" && BuyerRate > 0)
                                        {
                                            decimal MinRateLimit = 0m;
                                            decimal MiaxRateLimit = 0m;
                                            DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + PayoutCurrency + "'");
                                            if (objCurrency != null)
                                            {
                                                if (objCurrency.Rows.Count > 0)
                                                {
                                                    DataRow drCurrency = objCurrency.Rows[0];
                                                    MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                                    MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                                                }
                                                if (BuyerRate < MinRateLimit || BuyerRate > MiaxRateLimit)
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Buyer rate should be > 0<br>";
                                        }
                                        InsertCommand.Parameters.AddWithValue("@BuyerRate", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRate", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Buyer rate is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 20) // Buyer
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        BuyerName = Common.toString(ws.Row(r).Cell(i).Value);
                                        InsertCommand.Parameters.AddWithValue("@Buyer", BuyerName);
                                        AgentAccountCode = Common.getGLAccountCodeByPrefix(BuyerName);
                                        if (AgentAccountCode.Trim() != "")
                                        {
                                            AgentCurrencyCode = Common.getGLAccountCurrencyByCode(AgentAccountCode);
                                            if (AgentCurrencyCode != "")
                                            {
                                                if (!Common.isCurrencyAvailable(AgentCurrencyCode))
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += AgentCurrencyCode + " is not available or is inactive. Please check currency settings<br>";
                                                    AgentCurrencyCode = "";
                                                }
                                            }
                                            else
                                            {
                                                IsRowValid = false;
                                                rowErrMessage += "Currency not not setup for buyer<br>";
                                            }

                                            string sqlVendors = "select VendorID from Vendors where VendorPrefix ='" + Common.toString(BuyerName) + "'";
                                            string result = DBUtils.executeSqlGetSingle(sqlVendors);
                                            if (result.ToString().Trim() == "")
                                            {
                                                //** get data from GlChartofAccounts (prefix, title, code , country, currency) and save in Vendors
                                                string getInfo = "Select Prefix,AccountName, AccountCode,CountryISONumericCode,CurrencyISOCode from GLChartOfAccounts where AccountCode ='" + AgentAccountCode + "'";
                                                DataTable dtVendor = DBUtils.GetDataTable(getInfo);
                                                if (dtVendor != null)
                                                {
                                                    if (dtVendor.Rows.Count > 0)
                                                    {
                                                        DataRow drVendor = dtVendor.Rows[0];
                                                        string createVendor = "Insert into Vendors(VendorPrefix,VendorFirstName,AccountCode,VendorCountryISOCode,VendorCurrency,VendorCompany,Status) Values (@VendorPrefix,@VendorFirstName,@AccountCode,@VendorCountryISOCode,@VendorCurrency,@VendorCompany,@Status)";
                                                        SqlCommand Insertcommand = null;
                                                        Insertcommand = con.CreateCommand();
                                                        Insertcommand.Parameters.AddWithValue("@VendorPrefix", drVendor["Prefix"]);
                                                        Insertcommand.Parameters.AddWithValue("@VendorFirstName", drVendor["AccountName"]);
                                                        Insertcommand.Parameters.AddWithValue("@AccountCode", drVendor["AccountCode"]);
                                                        Insertcommand.Parameters.AddWithValue("@VendorCountryISOCode", drVendor["CountryISONumericCode"]);
                                                        Insertcommand.Parameters.AddWithValue("@VendorCurrency", drVendor["CurrencyISOCode"]);
                                                        Insertcommand.Parameters.AddWithValue("@VendorCompany", drVendor["AccountName"]);
                                                        Insertcommand.Parameters.AddWithValue("@Status", true);
                                                        Insertcommand.CommandText = createVendor;
                                                        Insertcommand.ExecuteNonQuery();
                                                        Insertcommand.Parameters.Clear();
                                                    }
                                                }

                                                //IsRowValid = false;
                                                //rowErrMessage += "Agent/customer does not exists in the system.<br>";                                        }
                                            }
                                            else
                                            {
                                                //IsRowValid = false;
                                                //rowErrMessage += "Invalid prefix. Buyer prefix not available in chart of account<br>";
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Buyer is not available in excel sheet<br>";
                                        }
                                    }
                                }
                                if (i == 21) // Buyer Rate SC
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRateSC", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRateSC", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Buyer Rate SC is not available in excel sheet<br>";
                                    }
                                }
                                if (i == 22) // Buyer Rate DC
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRateDC", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BuyerRateDC", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Buyer Rate DC is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 23) // BenePaymentMethod
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BenePaymentMethod", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@BenePaymentMethod", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Bene Payment Method is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 24) // SendingCountry
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@SendingCountry", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@SendingCountry", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Sending Country is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 25) // ReceivingCountry
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@ReceivingCountry", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@ReceivingCountry", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Receiving Country is not available in excel sheet<br>";
                                    }
                                }

                                if (i == 26) // Status
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Status", ws.Row(r).Cell(i).Value.ToString());
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@Status", "");
                                        //IsRowValid = false;
                                        //rowErrMessage += "Status is not available in excel sheet.<br>";
                                    }
                                }
                                if (i == 27) // Posting Date
                                {
                                    if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                                    {
                                        if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                        {
                                            //is valid date
                                            //is within fin year
                                            if (Common.isDate(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                            {
                                                if (Common.ValidateFiscalYearDate(Common.toDateTime(ws.Row(r).Cell(i).Value)))
                                                {
                                                    if (Common.ValidatePostingDate(Common.toString(ws.Row(r).Cell(i).Value)))
                                                    {
                                                        InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(ws.Row(r).Cell(i).Value));
                                                    }
                                                    else
                                                    {
                                                        InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                        IsRowValid = false;
                                                        rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                                    }
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                    IsRowValid = false;
                                                    rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                                                }

                                            }
                                            else
                                            {
                                                InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                IsRowValid = false;
                                                rowErrMessage += "Posting date is not valid.<br>";
                                            }

                                        }
                                        else
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                            IsRowValid = false;
                                            rowErrMessage += "Posting date is not available in excel sheet.<br>";
                                        }

                                    }
                                    else
                                    {
                                        if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                                        }
                                        else
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                            IsRowValid = false;
                                            rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                        }
                                    }
                                }
                            }

                            if (IsRowValid)
                            {
                                InsertCommand.Parameters.AddWithValue("@isRowValid", true);
                                InsertCommand.Parameters.AddWithValue("@cpUserRefNo", viewModel.cpUserRefNo);
                                InsertCommand.Parameters.AddWithValue("@RowErrMessage", rowErrMessage);
                            }
                            else
                            {
                                InsertCommand.Parameters.AddWithValue("@isRowValid", false);
                                InsertCommand.Parameters.AddWithValue("@RowErrMessage", rowErrMessage);
                                InsertCommand.Parameters.AddWithValue("@cpUserRefNo", viewModel.cpUserRefNo);
                            }

                            InsertCommand.Parameters.AddWithValue("@RowNumber", j);

                            InsertCommand.CommandText = sqlMyEntry;
                            InsertCommand.ExecuteNonQuery();
                            InsertCommand.Parameters.Clear();
                        }
                        rowErrMessage = "";
                        PayoutCurrency = "";
                        viewModel.isPosted = true;
                        j++;
                        //returnModel.isPosted = true;
                    }
                    con.Close();
                }
                else
                {
                    viewModel.ErrMessage = "Incomplete data in excel sheet.";
                }
            }
            catch (Exception ex)
            {
                viewModel.ErrMessage = "Unable to load data. Please try again. :" + ex.ToString();
            }
            return Json(viewModel, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ReadBranchPaidList([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM FileEntriesTemp /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                          Id,RowNumber,cpUserRefNo, TransactionDate, AgentName,PostingDate, Rate,PayoutCurrency,FCAmount,Payinamount,AdminCharges,AgentCharges,TRANID,CustomerId,CustomerName,FatherName,Recipient,Phone,Address,City,Code,BuyerRate,Buyer,BuyerRateSC,BuyerRateDC,PaymentNo,BenePaymentMethod,SendingCountry,ReceivingCountry,Status,isRowValid,RowErrMessage FROM FileEntriesTemp
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";
            //     const string selectQuery = @"SELECT
            //CheckListId, CheckTypeId, (select CheckType  from CheckTypes where CheckTypes.CheckTypeId = CheckList.CheckTypeId ) as CheckTypeName, CheckText, case when Status = 'true' then 'Active' else 'Inactive' end as StatusName, Status FROM CheckList";
            SqlBuilder builder = new SqlBuilder();
            var count = builder.AddTemplate(countQuery);

            int CurrentPage = 0; int CurrentPageSize = 1;
            if (request.PageSize > 0)
            {
                CurrentPageSize = request.PageSize;
            }
            if (request.PageSize > 0 && request.Page > 0)
            {
                CurrentPage = request.Page - 1;
            }
            var selector = builder.AddTemplate(selectQuery, new { PageIndex = CurrentPage, PageSize = CurrentPageSize });

            if (request.Filters != null && request.Filters.Any())
            {
                builder = Common.ApplyFilters(builder, request.Filters);
            }
            else
            {
                builder.Where("cpUserRefNo=''");
            }

            //cpUserRefNo
            //builder.Where("isHead='true'");

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("RowNumber");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);

            var rows = _dbcontext.Query<ImportPaidFileViewModel>(selector.RawSql, selector.Parameters);

            //  rows.Each(item => Console.WriteLine($"({ item.AccountCode}): { item.AccountName}
            //{                GetParentsString(items, item)}"));            
            //}
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult EditingInline_UpdateBranchPaid([DataSourceRequest] DataSourceRequest request, ImportPaidFileViewModel product)
        {
            if (product != null)
            {
                product = EditingCustom_UpdateBranchPaid(product);
            }

            return Json(new[] { product }.ToDataSourceResult(request, ModelState));
        }
        public ImportPaidFileViewModel EditingCustom_UpdateBranchPaid(ImportPaidFileViewModel pFileModel)
        {
            ApplicationUser Profile = new ApplicationUser();
            bool IsRowValid = true;
            string err_msgs = "";
            string UpdateSQL = "Update FileEntriesTemp Set ";
            UpdateSQL += "TransactionDate=@TransactionDate,PostingDate=@PostingDate,AgentName=@AgentName,AgentCharges=@AgentCharges,Rate=@Rate,PayoutCurrency=@PayoutCurrency,FCAmount=@FCAmount,Payinamount=@Payinamount, AdminCharges=@AdminCharges,TRANID=@TRANID,CustomerId=@CustomerId,CustomerName=@CustomerName,FatherName=@FatherName,Recipient=@Recipient,Phone=@Phone,Address=@Address,City=@City,Code=@Code,BuyerRate=@BuyerRate,Buyer=@Buyer,BuyerRateSC=@BuyerRateSC,BuyerRateDC=@BuyerRateDC,PaymentNo=@PaymentNo,BenePaymentMethod=@BenePaymentMethod,SendingCountry=@SendingCountry,ReceivingCountry=@ReceivingCountry,Status=@Status,isRowValid=@isRowValid,RowErrMessage=@RowErrMessage Where Id=@Id";
            using (SqlConnection con = Common.getConnection())
            {
                SqlCommand updateCommand = con.CreateCommand();
                try
                {
                    if (Common.toString(pFileModel.TransactionDate).Trim() != "")
                    {
                        updateCommand.Parameters.AddWithValue("@TransactionDate", pFileModel.TransactionDate);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Transaction date is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@TransactionDate", "");
                    }

                    if (Common.toString(pFileModel.PostingDate).Trim() != "")
                    {
                        if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                        {
                            if (Common.isDate(Common.toString(pFileModel.PostingDate)))
                            {
                                if (Common.ValidateFiscalYearDate(Common.toDateTime(pFileModel.PostingDate)))
                                {
                                    updateCommand.Parameters.AddWithValue("@PostingDate", pFileModel.PostingDate);
                                }
                                else
                                {
                                    updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                    IsRowValid = false;
                                    err_msgs += "Posting date is not between financial year start date and financial year end date.<br>";
                                }

                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Posting date is not valid.<br>";
                            }

                        }
                        else
                        {
                            if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", SysPrefs.PostingDate);
                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                            }
                        }

                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Posting date is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@PostingDate", "");
                    }

                    if (Common.toString(pFileModel.Buyer).Trim() != "")
                    {
                        string BuyerName = pFileModel.Buyer;
                        string AgentAccountCode = Common.getGLAccountCodeByPrefix(BuyerName);
                        if (AgentAccountCode.Trim() != "")
                        {
                            string AgentCurrencyCode = Common.getGLAccountCurrencyByCode(AgentAccountCode);
                            if (AgentCurrencyCode != "")
                            {
                                if (!Common.isCurrencyAvailable(AgentCurrencyCode))
                                {
                                    IsRowValid = false;
                                    err_msgs += AgentCurrencyCode + " is not available or is inactive. Please check currency settings<br>";
                                    AgentCurrencyCode = "";
                                }
                            }
                            else
                            {
                                IsRowValid = false;
                                err_msgs += "Currency not not setup for buyer<br>";
                            }

                            ////string sqlBuyers = "select VendorID from Vendors where VendorPrefix='" + Common.toString(BuyerName) + "'";
                            ////string result = DBUtils.executeSqlGetSingle(sqlBuyers);
                            ////if (result.ToString().Trim() == "")
                            ////{
                            ////    IsRowValid = false;
                            ////    err_msgs += "Buyer does not exists in the system.<br>";
                            ////}
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Invalid prefix. Buyer prefix not available in chart of account<br>";
                        }
                        updateCommand.Parameters.AddWithValue("@Buyer", pFileModel.Buyer);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Buyer is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@Buyer", "");
                    }

                    if (Common.toString(pFileModel.BuyerRate).Trim() != "")
                    {
                        decimal BuyerRate = pFileModel.BuyerRate;
                        if (Common.toString(pFileModel.PayoutCurrency).Trim() != "" && BuyerRate > 0)
                        {
                            decimal MinRateLimit = 0m;
                            decimal MiaxRateLimit = 0m;
                            DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + pFileModel.PayoutCurrency + "'");
                            if (objCurrency != null)
                            {
                                if (objCurrency.Rows.Count > 0)
                                {
                                    DataRow drCurrency = objCurrency.Rows[0];
                                    MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                    MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                                }
                                if (BuyerRate < MinRateLimit || BuyerRate > MiaxRateLimit)
                                {
                                    IsRowValid = false;
                                    err_msgs += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                                }
                            }
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Buyer rate should be > 0<br>";
                        }
                        updateCommand.Parameters.AddWithValue("@BuyerRate", pFileModel.BuyerRate);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Buyer rate is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@BuyerRate", "");
                    }

                    if (Common.toString(pFileModel.BuyerRateSC).Trim() != "")
                    {
                        updateCommand.Parameters.AddWithValue("@BuyerRateSC", pFileModel.BuyerRateSC);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Buyer Rate SC is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@BuyerRateSC", "");
                    }

                    if (Common.toString(pFileModel.BuyerRateDC).Trim() != "")
                    {
                        updateCommand.Parameters.AddWithValue("@BuyerRateDC", pFileModel.BuyerRateDC);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Buyer Rate DC is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@BuyerRateDC", "");
                    }

                    if (!string.IsNullOrEmpty(pFileModel.AgentName))
                    {
                        updateCommand.Parameters.AddWithValue("@AgentName", pFileModel.AgentName);
                        string accountcode = Common.getGLAccountCodeByPrefix(pFileModel.AgentName);
                        if (accountcode.Trim() != "")
                        {
                            //string sqlCustomers = "select CustomerID from Customers where CustomerPrefix='" + Common.toString(pFileModel.AgentName) + "'";
                            //string result = DBUtils.executeSqlGetSingle(sqlCustomers);
                            //if (result.ToString().Trim() == "")
                            //{
                            //    IsRowValid = false;
                            //    err_msgs += "Agent/customer does not exists in the system.<br>";
                            //}
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Invalid prefix. Agent prefix not available in chart of account<br>";
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Agent Name is not available in editable grid<br>";
                        updateCommand.Parameters.AddWithValue("@AgentName", "");
                    }
                    // Agent_rate
                    if (pFileModel.Rate > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Rate", pFileModel.Rate);
                        //if (!string.IsNullOrEmpty(pFileModel.PayoutCurrency) )
                        //{
                        //    decimal MinRateLimit = 0m;
                        //    decimal MiaxRateLimit = 0m;
                        //    DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + pFileModel.PayoutCurrency + "'.<br>");
                        //    if (objCurrency != null)
                        //    {
                        //        if (objCurrency.Rows.Count > 0)
                        //        {
                        //            DataRow drCurrency = objCurrency.Rows[0];
                        //            MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                        //            MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                        //        }
                        //        if (pFileModel.Rate < MinRateLimit || pFileModel.Rate > MiaxRateLimit)
                        //        {
                        //            IsRowValid = false;
                        //            err_msgs += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                        //        }
                        //    }
                        //}
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Rate is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@Rate", "");
                    }
                    // PayOut_CCY
                    if (!string.IsNullOrEmpty(pFileModel.PayoutCurrency))
                    {
                        updateCommand.Parameters.AddWithValue("@PayoutCurrency", pFileModel.PayoutCurrency);
                        string PayoutCurrency = pFileModel.PayoutCurrency;
                        if (Common.toString(PayoutCurrency).Trim() != "")
                        {
                            if (Common.isCurrencyAvailable(PayoutCurrency))
                            {
                                string accountcode = Common.getHoldAccountByCurrency(PayoutCurrency, "1");
                                if (accountcode == "")
                                {
                                    AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(PayoutCurrency, "1");
                                    if (!response.isAdded)
                                    {
                                        IsRowValid = false;
                                        err_msgs += PayoutCurrency + " Unpaid Account is not available and unable to create it. Please check currency settings.<br>";
                                    }
                                }
                            }
                            else
                            {
                                IsRowValid = false;
                                err_msgs += PayoutCurrency + " is not available or is inactive. Please check currency settings.<br>";
                            }
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Payout currency is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@PayoutCurrency", "");
                    }

                    // Amount
                    if (pFileModel.FCAmount > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@FCAmount", pFileModel.FCAmount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "F/C amount is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@FCAmount", "0");
                    }
                    // Payin_amount
                    if (pFileModel.Payinamount >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Payinamount", pFileModel.Payinamount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Payin amount is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@Payinamount", "0");
                    }
                    /// Admin Charges
                    if (pFileModel.AdminCharges >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@AdminCharges", pFileModel.AdminCharges);
                        ////if (pFileModel.AdminCharges > 0 && Common.toString(pFileModel.PayoutCurrency) != "")
                        ////{
                        ////    string accountcode = Common.getHoldAccountByCurrency(pFileModel.PayoutCurrency, "2");
                        ////    if (accountcode == "")
                        ////    {
                        ////        AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(pFileModel.PayoutCurrency, "2");
                        ////        if (!response.isAdded)
                        ////        {
                        ////            IsRowValid = false;
                        ////            err_msgs += pFileModel.PayoutCurrency + " Un-earned admin charges account is not available and unable to create it. Please check currency settings.<br>";
                        ////        }
                        ////    }
                        ////}
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@AdminCharges", "0");
                    }

                    ////     Agent Charges 
                    if (pFileModel.AgentCharges >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@AgentCharges", pFileModel.AgentCharges);
                        ////if (pFileModel.AgentCharges > 0 && Common.toString(pFileModel.PayoutCurrency) != "")
                        ////{
                        ////    string accountcode = Common.getHoldAccountByCurrency(pFileModel.PayoutCurrency, "3");
                        ////    if (accountcode == "")
                        ////    {
                        ////        AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(pFileModel.PayoutCurrency, "3");
                        ////        if (!response.isAdded)
                        ////        {
                        ////            IsRowValid = false;
                        ////            err_msgs += pFileModel.PayoutCurrency + " Un-earned agent charges account is not available and unable to create it. Please check currency settings.<br>";
                        ////        }
                        ////    }
                        ////}
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@AgentCharges", "0");
                    }
                    /// Phone
                    if (!string.IsNullOrEmpty(pFileModel.Phone))
                    {
                        updateCommand.Parameters.AddWithValue("@Phone", pFileModel.Phone);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Phone is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Phone", "");
                    }
                    /// FatherName
                    if (!string.IsNullOrEmpty(pFileModel.FatherName))
                    {
                        updateCommand.Parameters.AddWithValue("@FatherName", pFileModel.FatherName);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "FatherName is not available in excel sheet;";
                        updateCommand.Parameters.AddWithValue("@FatherName", "");
                    }
                    /// Address
                    if (!string.IsNullOrEmpty(pFileModel.Address))
                    {
                        updateCommand.Parameters.AddWithValue("@Address", pFileModel.Address);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Address is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Address", "");
                    }
                    /// City
                    if (!string.IsNullOrEmpty(pFileModel.City))
                    {
                        updateCommand.Parameters.AddWithValue("@City", pFileModel.City);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "City is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@City", "");
                    }
                    /// Code
                    if (!string.IsNullOrEmpty(pFileModel.Code))
                    {
                        updateCommand.Parameters.AddWithValue("@Code", pFileModel.Code);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Code is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Code", "");
                    }
                    /// PaymentNo
                    if (!string.IsNullOrEmpty(pFileModel.PaymentNo))
                    {
                        updateCommand.Parameters.AddWithValue("@PaymentNo", pFileModel.PaymentNo);
                        if (!string.IsNullOrEmpty(Common.toString(pFileModel.PaymentNo).Trim()))
                        {
                            string sqlPaymentNo = "select PaymentNo from CustomerTransactions where PaymentNo='" + Common.toString(pFileModel.PaymentNo).Trim() + "'";
                            string result = DBUtils.executeSqlGetSingle(sqlPaymentNo);
                            if (Common.toString(result).Trim() == "")
                            {
                                IsRowValid = false;
                                err_msgs += "Can not find this payment # (" + Common.toString(pFileModel.PaymentNo) + ") in unpaid/hold transaction list.<br>";
                            }
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Payment No is not available in editable grid<br>";
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "PaymentNo is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@PaymentNo", "");
                    }
                    /// BenePaymentMethod
                    if (!string.IsNullOrEmpty(pFileModel.BenePaymentMethod))
                    {
                        updateCommand.Parameters.AddWithValue("@BenePaymentMethod", pFileModel.BenePaymentMethod);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Bene Payment Method is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@BenePaymentMethod", "");
                    }
                    /// SendingCountry
                    if (!string.IsNullOrEmpty(pFileModel.SendingCountry))
                    {
                        updateCommand.Parameters.AddWithValue("@SendingCountry", pFileModel.SendingCountry);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Sending Country is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@SendingCountry", "");
                    }
                    /// ReceivingCountry
                    if (!string.IsNullOrEmpty(pFileModel.ReceivingCountry))
                    {
                        updateCommand.Parameters.AddWithValue("@ReceivingCountry", pFileModel.ReceivingCountry);
                    }
                    else
                    {
                        //   IsRowValid = false;
                        //err_msgs += "Receiving Country is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@ReceivingCountry", "");
                    }
                    /// Tr_No
                    if (pFileModel.TRANID > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@TRANID", pFileModel.TRANID);
                        if (!string.IsNullOrEmpty(Common.toString(pFileModel.TRANID).Trim()))
                        {
                            string sqlTranx = "select TransactionID from CustomerTransactions where TransactionID='" + pFileModel.TRANID + "'";
                            string result = DBUtils.executeSqlGetSingle(sqlTranx);
                            if (Common.toString(result).Trim() == "")
                            {
                                IsRowValid = false;
                                err_msgs += "Can not find this transaction (" + pFileModel.TRANID + ") in unpaid/hold transaction list.<br>";
                            }
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Transaction Id is not available in editable grid<br>";
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Transaction id is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@TRANID", "");
                    }
                    /// Customer_id
                    if (pFileModel.CustomerId > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerId", pFileModel.CustomerId);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Customer id is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@CustomerId", "");
                    }
                    /// Customer_full_name
                    if (!string.IsNullOrEmpty(pFileModel.CustomerName))
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerName", pFileModel.CustomerName);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Customer name is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@CustomerName", "");
                    }
                    /// Beneficiary_full_name
                    if (!string.IsNullOrEmpty(pFileModel.Recipient))
                    {
                        updateCommand.Parameters.AddWithValue("@Recipient", pFileModel.Recipient);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Recipient is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Recipient", "");
                    }

                    if (!string.IsNullOrEmpty(pFileModel.Status))
                    {
                        updateCommand.Parameters.AddWithValue("@Status", pFileModel.Status);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "Status is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Status", "");
                    }

                    if (IsRowValid)
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", true);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                        updateCommand.Parameters.AddWithValue("@RowErrMessage", err_msgs);
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", false);
                        updateCommand.Parameters.AddWithValue("@RowErrMessage", err_msgs);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                    }
                    updateCommand.CommandText = UpdateSQL;
                    updateCommand.ExecuteNonQuery();
                    updateCommand.Parameters.Clear();

                }
                catch (Exception ex)
                {

                }
                finally
                {
                    con.Close();
                }
            }
            pFileModel.isRowValid = IsRowValid;
            pFileModel.RowErrMessage = err_msgs;
            return pFileModel;
        }

        public ActionResult PostBranchPaidData(ImportPaidFileViewModel pPostingModel)
        {
            ErrorMessages objErrorMessage = null;
            ////  string myConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
            ////  string myConnectionString = BaseModel.getConnString();
            ApplicationUser profile = ApplicationUser.GetUserProfile();
            DataSet DS_lstGLEntriesTemp = DBUtils.GetDataSet("select * from FileEntriesTemp where cpUserRefNo='" + pPostingModel.cpUserRefNo + "'");
            DataTable lstGLEntriesTemp = DS_lstGLEntriesTemp.Tables[0];
            if (lstGLEntriesTemp != null)
            {

                if (lstGLEntriesTemp.Rows.Count > 0)
                {
                    bool isMainValid = true;
                    foreach (DataRow objGLEntriesTemp in lstGLEntriesTemp.Rows)
                    {
                        if (Common.toBool(objGLEntriesTemp["isRowValid"]) == false)
                        {
                            isMainValid = false;
                        }
                    }
                    if (isMainValid)
                    {

                        objErrorMessage = PostingUtils.SaveBranchPaidTransactionBridge("22PAIDENTRY", lstGLEntriesTemp, Profile.Id.ToString());
                        string postingErr = "";
                        if (objErrorMessage.IsFailed)
                        {
                            for (int i = 0; i < objErrorMessage.ErrorMessage.Count; i++)
                            {
                                postingErr = postingErr + "<br>" + objErrorMessage.ErrorMessage[i].Message;
                            }
                            pPostingModel.ErrMessage = Common.GetAlertMessage(1, postingErr);
                        }
                        if (objErrorMessage.ErrorMessage.Count > 0)
                        {
                            pPostingModel.ErrMessage = objErrorMessage.ErrorMessage[0].Message.ToString();
                        }
                        else
                        {
                            pPostingModel.ErrMessage = Common.GetAlertMessage(0, "Paid transaction from excel is successfully added <br> <br> " + objErrorMessage.ErrorMessage[0].Message + "  <br> <br> <a href ='/Inquiries/GLInquiry'>Click here</a>");
                        }
                    }
                    else
                    {
                        pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Please check your data and try again");
                    }
                }
            }
            else
            {
                pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Data not found");
            }
            HoldFilePostingViewModel pModel = new HoldFilePostingViewModel();
            pModel.ErrMessage = pPostingModel.ErrMessage;

            return View("~/Views/ImportExcel/Message.cshtml", pModel);
        }
        #endregion

        #region Import Cancelled File
        public ActionResult CancelledFile()
        {
            DBUtils.ExecuteSQL("Delete from FileEntriesTemp where (cpUserRefNo = '') or (cpUserRefNo is null) or (cpUserRefNo='0')");
            ImportCancelledFileViewModel postingViewModel = new ImportCancelledFileViewModel();
            Guid guid = Guid.NewGuid();
            postingViewModel.cpUserRefNo = guid.ToString();
            if (string.IsNullOrEmpty(SysPrefs.TransactionAdminCommissionAccount))
            {
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Unable to find transaction admin commission account. Please update preferences.");
            }
            if (string.IsNullOrEmpty(SysPrefs.DefaultCurrency))
            {
                postingViewModel.ErrMessage += Common.GetAlertMessage(1, "Unable to find default curreny. Please update preferences.");
            }
            if (string.IsNullOrEmpty(SysPrefs.PostingDate))
            {
                postingViewModel.ErrMessage += Common.GetAlertMessage(1, "Unable to find posting date. Please update preferences.");
            }
            //if (string.IsNullOrEmpty(SysPrefs.UnEarnedServiceFeeAccount))
            //{
            //    postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Unable to find Un-Earned Service fee account.");
            //}
            return View("~/Views/ImportExcel/CancelledFile.cshtml", postingViewModel);
        }

        [HttpPost]
        public ActionResult CancelledFileUploaded(ImportCancelledFileViewModel pImportCancelledFileViewModel, HttpPostedFileBase file)
        {
            try
            {
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                IEnumerable<Sheets> pcheckListViewModel = new List<Sheets>();
                var fileName = ""; var path = "";
                pImportCancelledFileViewModel.isFileImported = false;
                if (file != null && file.ContentLength > 0)
                {
                    var supportedTypes = new[] { "xlsx" };
                    var fileExt = Path.GetExtension(file.FileName).Substring(1);
                    if (!supportedTypes.Contains(fileExt))
                    {
                        pImportCancelledFileViewModel.ErrMessage = Common.GetAlertMessage(1, "File extension is invalid -  only xlsx files are allowed.");
                        pImportCancelledFileViewModel.isError = true;
                    }
                    else
                    {
                        fileName = Path.GetFileName(file.FileName);
                        ////fileName = Common.getRandomDigit(6) + "-" + fileName;
                        //// path = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/"), fileName);
                        fileName = "Cancelled-" + DateTime.Now.ToString("MMddyyyyHHmmssffff") + ".xlsx";
                        path = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/UpTemp/"), fileName);
                        file.SaveAs(path);
                        pImportCancelledFileViewModel.isFileImported = true;

                        pImportCancelledFileViewModel.ExcelFilePath = path;

                        if (path != null && path != "")
                        {
                            int Count = 0;
                            var workbook = new XLWorkbook(path);
                            foreach (var worksheet in workbook.Worksheets)
                            {
                                var sheetname = worksheet.ToString().Split('!');
                                var sheetname1 = sheetname[0].ToString();
                                sheetname1 = RemoveSpecialCharacters(sheetname1);

                                Sheets pSheets = new Sheets();
                                pSheets.Name = Common.toString(sheetname1);
                                //pSheets.Name = Common.toString(fileName);
                                pSheets.Name = pSheets.Name.Replace("'", "");
                                pSheets.Value = Count.ToString();
                                pcheckListViewModel = (new[] { pSheets }).Concat(pcheckListViewModel);
                                Count++;
                            }

                            pImportCancelledFileViewModel.sheets = pcheckListViewModel.ToList();

                        }
                        else
                        {
                            pImportCancelledFileViewModel.ErrMessage = "file not uploaded. please try again.";
                            pImportCancelledFileViewModel.isError = true;
                        }
                    }
                }
                else
                {
                    pImportCancelledFileViewModel.ErrMessage = "file not uploaded. please try again.";
                    pImportCancelledFileViewModel.isError = true;
                }
            }
            catch (Exception ex)
            {
                pImportCancelledFileViewModel.ErrMessage = "An error occured: " + ex.ToString();
                pImportCancelledFileViewModel.isError = true;
            }
            return View("~/Views/ImportExcel/CancelledFile.cshtml", pImportCancelledFileViewModel);
        }
        public ActionResult ShowCancelledExcelfiledata(ImportCancelledFileViewModel viewModel)
        {
            string sqlMyEntry = "";
            string autoSetPostingDate = Common.GetSysSettings("AutoSetPostingDate");
            //HoldFilePostingViewModel pHoldFilePostingViewModel = new HoldFilePostingViewModel();
            try
            {
                string savepath = viewModel.ExcelFilePath;
                var workbook = new XLWorkbook(savepath);
                IXLWorksheet ws;
                string PayoutCurrency = "";
                bool IsRowValid = true;
                string rowErrMessage = "";
                string sheetname = viewModel.Sheet;
                workbook.TryGetWorksheet(Common.toString(sheetname), out ws);
                var datarange = ws.RangeUsed();
                int TotalRowsinExcelSheet = datarange.RowCount();
                int TotalColinExcelSheet = datarange.ColumnCount();
                int TotalFileColumns = 23;
                //if (SysPrefs.ImportFileHasPostingDate == "1")
                //{
                //    TotalFileColumns = 11;
                //}
                if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                {
                    TotalFileColumns = 24;
                }
                if (TotalColinExcelSheet == TotalFileColumns)
                {
                    int row_number = 0;
                    if (viewModel.HasHeader)
                    {
                        row_number = 2;
                    }
                    else
                    {
                        row_number = 1;
                    }
                    int row_counter = 0;
                    int j = 1;
                    SqlCommand InsertCommand = null;
                    SqlConnection con = Common.getConnection();
                    for (int r = row_number; r <= TotalRowsinExcelSheet; r++)
                    {
                        sqlMyEntry = "Insert into FileEntriesTemp (RowNumber,TransactionDate,AgentName,Rate,PayoutCurrency,FCAmount,Payinamount, AdminCharges,AgentCharges,TRANID,CustomerId,CustomerName,FatherName,Recipient,Phone,Address,City,Code,PaymentNo,BenePaymentMethod,SendingCountry,ReceivingCountry,Status, isRowValid, RowErrMessage,cpUserRefNo,Charges,PostingDate) Values (@RowNumber,@TransactionDate,@AgentName,@Rate,@PayoutCurrency,@FCAmount,@Payinamount, @AdminCharges,@AgentCharges,@TRANID,@CustomerId,@CustomerName,@FatherName,@Recipient,@Phone,@Address,@City,@Code,@PaymentNo,@BenePaymentMethod,@SendingCountry,@ReceivingCountry,@Status,@isRowValid,@RowErrMessage,@cpUserRefNo,@Charges,@PostingDate)";
                        string AgentName = "";
                        string AgentAccountCode = "";
                        string AgentCurrencyCode = "";
                        InsertCommand = con.CreateCommand();
                        //if (ws.Row(r).Cell(22).Value.ToString() == "Refunded")
                        //{
                        if (Common.GetSysSettings("ImportFileHasPostingDate") == "0")
                        {
                            if (Common.ValidateFiscalYearDate(Common.toDateTime(SysPrefs.PostingDate)))
                            {
                                InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                            }
                            else
                            {
                                InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                            }
                        }
                        for (int i = 1; i <= TotalColinExcelSheet; i++)
                        {
                            if (i == 1) // Agent Name
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    AgentName = Common.toString(ws.Row(r).Cell(i).Value);
                                    InsertCommand.Parameters.AddWithValue("@AgentName", AgentName);
                                    AgentAccountCode = Common.getGLAccountCodeByPrefix(AgentName);
                                    if (AgentAccountCode.Trim() != "")
                                    {
                                        AgentCurrencyCode = Common.getGLAccountCurrencyByCode(AgentAccountCode);
                                        if (AgentCurrencyCode != "")
                                        {
                                            if (!Common.isCurrencyAvailable(AgentCurrencyCode))
                                            {
                                                IsRowValid = false;
                                                rowErrMessage += AgentCurrencyCode + " is not available or is inactive. Please check currency settings<br>";
                                                AgentCurrencyCode = "";
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += "Currency not not setup for agent<br>";
                                        }
                                    }
                                    else
                                    {
                                        IsRowValid = false;
                                        rowErrMessage += "Invalid prefix. Agent prefix not available in chart of account<br>";
                                    }
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@AgentName", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Agent name is not available in excel sheet<br>";
                                }
                            }

                            if (i == 2) // Customer_id
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    int CustomerId = Common.toInt(ws.Row(r).Cell(i).Value);
                                    InsertCommand.Parameters.AddWithValue("@CustomerId", CustomerId);
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@CustomerId", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "Customer id is not available in excel sheet<br>";
                                }
                            }

                            if (i == 3) // Customer_Name
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@CustomerName", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@CustomerName", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Customer name is not available in excel sheet<br>";
                                }
                            }

                            if (i == 4) // TRAN ID
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@TRANID", Common.toInt(ws.Row(r).Cell(i).Value));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@TRANID", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "Transaction Id is not available in excel sheet<br>";
                                }
                            }

                            if (i == 5) // Beneficiary_full_name
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Recipient", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Recipient", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Recipient is not available in excel sheet<br> ";
                                }
                            }


                            if (i == 6) // FatherName
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@FatherName", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@FatherName", "");
                                    //IsRowValid = false;
                                    //rowErrMessage += "FatherName is not available in excel sheet; ";
                                }
                            }
                            if (i == 7) // Phone
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Phone", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Phone", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Phone is not available in excel sheet<br> ";
                                }
                            }

                            if (i == 8) // Address
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Address", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Address", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Address is not available in excel sheet<br>";
                                }
                            }


                            if (i == 9) // City
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@City", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@City", "");
                                    IsRowValid = false;
                                    rowErrMessage += "City is not available in excel sheet<br>";
                                }
                            }

                            if (i == 10) // Code
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Code", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Code", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Code is not available in excel sheet<br> ";
                                }
                            }

                            if (i == 11) // PayOut_CCY
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    PayoutCurrency = Common.toString(ws.Row(r).Cell(i).Value);
                                    if (Common.toString(PayoutCurrency).Trim() != "")
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PayoutCurrency", PayoutCurrency);
                                        if (Common.isCurrencyAvailable(PayoutCurrency))
                                        {
                                            string accountcode = Common.getHoldAccountByCurrency(PayoutCurrency, "1");
                                            if (accountcode == "")
                                            {
                                                AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(PayoutCurrency, "1");
                                                if (!response.isAdded)
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += ws.Row(r).Cell(i).Value.ToString() + " Unpaid Account is not available and unable to create it. Please check currency settings.<br>";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            IsRowValid = false;
                                            rowErrMessage += PayoutCurrency + " is not available or is inactive. Please check currency settings<br>";
                                        }
                                    }
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@PayoutCurrency", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Payout currency is not available in excel sheet<br>";
                                }
                            }

                            if (i == 12) // PayOut_Amount
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@FCAmount", Common.toDecimal(ws.Row(r).Cell(i).Value));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@FCAmount", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "F/C amount is not available in excel sheet<br>";
                                }
                            }
                            if (i == 13) // Agent_rate
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    decimal Rate = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                    //if (Common.toString(PayoutCurrency).Trim() != "" && Rate > 0)
                                    //{
                                    //    decimal MinRateLimit = 0m;
                                    //    decimal MiaxRateLimit = 0m;
                                    //    DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + PayoutCurrency + "'");
                                    //    if (objCurrency != null)
                                    //    {
                                    //        if (objCurrency.Rows.Count > 0)
                                    //        {
                                    //            DataRow drCurrency = objCurrency.Rows[0];
                                    //            MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                    //            MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                                    //        }
                                    //        if (Rate < MinRateLimit || Rate > MiaxRateLimit)
                                    //        {
                                    //            IsRowValid = false;
                                    //            rowErrMessage += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                                    //        }
                                    //    }
                                    //}
                                    InsertCommand.Parameters.AddWithValue("@Rate", Rate);
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Rate", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "Rate is not available in excel sheet<br>";
                                }
                            }
                            if (i == 14) // Payin_amount
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Payinamount", Common.toDecimal(ws.Row(r).Cell(i).Value));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Payinamount", 0);
                                    IsRowValid = false;
                                    rowErrMessage += "Pay in amount is not available in excel sheet<br>";
                                }
                            }

                            if (i == 15) // PaymentNo
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@PaymentNo", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@PaymentNo", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Payment No is not available in excel sheet<br>";
                                }
                            }

                            if (i == 16) // TransactionDate
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@TransactionDate", Common.toDateTime(ws.Row(r).Cell(i).Value));
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@TransactionDate", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Transaction date is not available in excel sheet<br>";
                                }
                            }
                            if (i == 17) // Admin_Charges
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    decimal AdminCharges = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                    InsertCommand.Parameters.AddWithValue("@AdminCharges", AdminCharges);
                                    if (AdminCharges > 0)
                                    {
                                        if (AgentAccountCode != "" && AgentCurrencyCode != "")
                                        {
                                            string accountcode = Common.getHoldAccountByCurrency(AgentCurrencyCode, "2");
                                            if (accountcode == "")
                                            {
                                                AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(AgentCurrencyCode, "2");
                                                if (!response.isAdded)
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += AgentCurrencyCode + " Un-earned admin charges account is not available and unable to create it. Please check currency settings.<br>";
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@AdminCharges", 0);
                                    //IsRowValid = false;
                                    //rowErrMessage += "Admin_Charges is not available in excel sheet;";
                                }
                            }
                            if (i == 18) // Agent_Charges
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    decimal AgentCharges = Common.toDecimal(ws.Row(r).Cell(i).Value);
                                    InsertCommand.Parameters.AddWithValue("@AgentCharges", AgentCharges);
                                    if (AgentCharges > 0)
                                    {
                                        if (AgentAccountCode != "" && AgentCurrencyCode != "")
                                        {
                                            string accountcode = Common.getHoldAccountByCurrency(AgentCurrencyCode, "3");
                                            if (accountcode == "")
                                            {
                                                AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(AgentCurrencyCode, "3");
                                                if (!response.isAdded)
                                                {
                                                    IsRowValid = false;
                                                    rowErrMessage += AgentCurrencyCode + " Un-earned agent charges account is not available and unable to create it. Please check currency settings.<br>";
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@AgentCharges", 0);
                                }
                            }

                            if (i == 19) // BenePaymentMethod
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@BenePaymentMethod", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@BenePaymentMethod", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Bene Payment Method is not available in excel sheet<br>";
                                }
                            }

                            if (i == 20) // SendingCountry
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@SendingCountry", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@SendingCountry", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Sending Country is not available in excel sheet<br>";
                                }
                            }

                            if (i == 21) // ReceivingCountry
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@ReceivingCountry", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@ReceivingCountry", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Receiving Country is not available in excel sheet<br>";
                                }
                            }
                            if (i == 22) // ReceivingCountry
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Status", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Status", "");
                                    IsRowValid = false;
                                    rowErrMessage += "Status is not available in excel sheet<br>";
                                }
                            }
                            if (i == 23) // ReceivingCountry
                            {
                                if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                {
                                    InsertCommand.Parameters.AddWithValue("@Charges", ws.Row(r).Cell(i).Value.ToString());
                                }
                                else
                                {
                                    InsertCommand.Parameters.AddWithValue("@Charges", "0");
                                    IsRowValid = false;
                                    rowErrMessage += "Charges is not available in excel sheet<br>";
                                }
                            }
                            if (i == 24) // Posting Date
                            {
                                if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                                {
                                    if (!string.IsNullOrEmpty(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                    {
                                        //is valid date
                                        //is within fin year
                                        if (Common.isDate(Common.toString(ws.Row(r).Cell(i).Value).Trim()))
                                        {
                                            if (Common.ValidateFiscalYearDate(Common.toDateTime(ws.Row(r).Cell(i).Value)))
                                            {
                                                if (Common.ValidatePostingDate(Common.toString(ws.Row(r).Cell(i).Value)))
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(ws.Row(r).Cell(i).Value));
                                                }
                                                else
                                                {
                                                    InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                    IsRowValid = false;
                                                    rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                                }
                                            }
                                            else
                                            {
                                                InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                                IsRowValid = false;
                                                rowErrMessage += "Posting date is not between financial year start date and financial year end date.<br>";
                                            }

                                        }
                                        else
                                        {
                                            InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                            IsRowValid = false;
                                            rowErrMessage += "Posting date is not valid.<br>";
                                        }

                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Posting date is not available in excel sheet.<br>";
                                    }

                                }
                                else
                                {

                                    if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                                    }
                                    else
                                    {
                                        InsertCommand.Parameters.AddWithValue("@PostingDate", "");
                                        IsRowValid = false;
                                        rowErrMessage += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                    }
                                }
                            }

                        }

                        if (IsRowValid)
                        {
                            InsertCommand.Parameters.AddWithValue("@isRowValid", true);
                            InsertCommand.Parameters.AddWithValue("@cpUserRefNo", viewModel.cpUserRefNo);
                            InsertCommand.Parameters.AddWithValue("@RowErrMessage", rowErrMessage);
                        }
                        else
                        {
                            InsertCommand.Parameters.AddWithValue("@isRowValid", false);
                            InsertCommand.Parameters.AddWithValue("@RowErrMessage", rowErrMessage);
                            InsertCommand.Parameters.AddWithValue("@cpUserRefNo", viewModel.cpUserRefNo);
                        }
                        InsertCommand.Parameters.AddWithValue("@RowNumber", j);
                        InsertCommand.CommandText = sqlMyEntry;
                        InsertCommand.ExecuteNonQuery();
                        InsertCommand.Parameters.Clear();
                        rowErrMessage = "";
                        PayoutCurrency = "";
                        viewModel.isPosted = true;
                        j++;
                        //returnModel.isPosted = true;
                    }
                    con.Close();
                }
                else
                {
                    viewModel.ErrMessage = "Incomplete data in excel sheet.";
                }

            }
            catch (Exception ex)
            {
                viewModel.ErrMessage = "Unable to load data Please try Again.:" + ex.Message.ToString();
            }
            return Json(viewModel, JsonRequestBehavior.AllowGet);
        }
        public ActionResult CheckList_ReadCancelled([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM FileEntriesTemp /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                          Id,RowNumber,cpUserRefNo, TransactionDate,PostingDate, AgentName, Rate,PayoutCurrency,FCAmount,Payinamount,AdminCharges,AgentCharges,TRANID,CustomerId,CustomerName,FatherName,Recipient,Phone,Address,City,Code,PaymentNo,BenePaymentMethod,SendingCountry,ReceivingCountry,Status,Charges,isRowValid,RowErrMessage FROM FileEntriesTemp
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";
            //     const string selectQuery = @"SELECT
            //CheckListId, CheckTypeId, (select CheckType  from CheckTypes where CheckTypes.CheckTypeId = CheckList.CheckTypeId ) as CheckTypeName, CheckText, case when Status = 'true' then 'Active' else 'Inactive' end as StatusName, Status FROM CheckList";
            SqlBuilder builder = new SqlBuilder();
            var count = builder.AddTemplate(countQuery);

            int CurrentPage = 0; int CurrentPageSize = 1;
            if (request.PageSize > 0)
            {
                CurrentPageSize = request.PageSize;
            }
            if (request.PageSize > 0 && request.Page > 0)
            {
                CurrentPage = request.Page - 1;
            }
            var selector = builder.AddTemplate(selectQuery, new { PageIndex = CurrentPage, PageSize = CurrentPageSize });

            if (request.Filters != null && request.Filters.Any())
            {
                builder = Common.ApplyFilters(builder, request.Filters);
            }
            else
            {
                builder.Where("cpUserRefNo=''");
            }
            //cpUserRefNo
            //builder.Where("isHead='true'");

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("RowNumber");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);
            var rows = _dbcontext.Query<ImportCancelledFileViewModel>(selector.RawSql, selector.Parameters);
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult EditingInline_UpdateCancelled([DataSourceRequest] DataSourceRequest request, ImportCancelledFileViewModel product)
        {
            if (product != null)
            {
                product = EditingCustom_UpdateCancelled(product);
            }

            return Json(new[] { product }.ToDataSourceResult(request, ModelState));
        }
        public ImportCancelledFileViewModel EditingCustom_UpdateCancelled(ImportCancelledFileViewModel pFileModel)
        {
            ApplicationUser Profile = new ApplicationUser();
            bool IsRowValid = true;
            string err_msgs = "";
            string UpdateSQL = "Update FileEntriesTemp Set ";
            UpdateSQL += "TransactionDate=@TransactionDate,PostingDate=@PostingDate,AgentName=@AgentName,AgentCharges=@AgentCharges,Rate=@Rate,PayoutCurrency=@PayoutCurrency,FCAmount=@FCAmount,Payinamount=@Payinamount, AdminCharges=@AdminCharges,TRANID=@TRANID,CustomerId=@CustomerId,CustomerName=@CustomerName,FatherName=@FatherName,Recipient=@Recipient,Phone=@Phone,Address=@Address,City=@City,Code=@Code,PaymentNo=@PaymentNo,BenePaymentMethod=@BenePaymentMethod,SendingCountry=@SendingCountry,ReceivingCountry=@ReceivingCountry,Status=@Status,isRowValid=@isRowValid,RowErrMessage=@RowErrMessage,Charges=@Charges Where Id=@Id";
            using (SqlConnection con = Common.getConnection())
            {
                SqlCommand updateCommand = con.CreateCommand();
                try
                {
                    if (Common.toString(pFileModel.TransactionDate).Trim() != "")
                    {
                        updateCommand.Parameters.AddWithValue("@TransactionDate", pFileModel.TransactionDate);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Transaction date is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@TransactionDate", "");
                    }

                    if (Common.toString(pFileModel.PostingDate).Trim() != "")
                    {
                        if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
                        {
                            if (Common.isDate(Common.toString(pFileModel.PostingDate)))
                            {
                                if (Common.ValidateFiscalYearDate(Common.toDateTime(pFileModel.PostingDate)))
                                {
                                    if (Common.ValidatePostingDate(Common.toString(pFileModel.PostingDate)))
                                    {
                                        updateCommand.Parameters.AddWithValue("@PostingDate", pFileModel.PostingDate);
                                    }
                                    else
                                    {
                                        updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                        IsRowValid = false;
                                        err_msgs += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                                    }
                                }
                                else
                                {
                                    updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                    IsRowValid = false;
                                    err_msgs += "Posting date is not between financial year start date and financial year end date.<br>";
                                }

                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Posting date is not valid.<br>";
                            }

                        }
                        else
                        {
                            if (Common.ValidatePostingDate(Common.toString(SysPrefs.PostingDate)))
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", SysPrefs.PostingDate);
                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@PostingDate", "");
                                IsRowValid = false;
                                err_msgs += "Invalid Posting date . Posting Date should be greater or equal than [" + SysPrefs.PostingStartDate + "].<br>";
                            }

                        }

                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Posting date is not available in editable grid.<br>";
                        updateCommand.Parameters.AddWithValue("@PostingDate", "");
                    }

                    if (!string.IsNullOrEmpty(pFileModel.AgentName))
                    {
                        updateCommand.Parameters.AddWithValue("@AgentName", pFileModel.AgentName);
                        string accountcode = Common.getGLAccountCodeByPrefix(pFileModel.AgentName);
                        if (accountcode.Trim() != "")
                        {
                            string sqlCustomers = "select CustomerID from Customers where CustomerPrefix='" + Common.toString(pFileModel.AgentName) + "'";
                            string result = DBUtils.executeSqlGetSingle(sqlCustomers);
                            if (result.ToString().Trim() == "")
                            {
                                IsRowValid = false;
                                err_msgs += "Agent/customer does not exists in the system.<br>";
                            }
                        }
                        else
                        {
                            IsRowValid = false;
                            err_msgs += "Invalid prefix. Agent prefix not available in chart of account<br>";
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Agent Name is not available in excel sheet<br>";
                        updateCommand.Parameters.AddWithValue("@AgentName", "");
                    }
                    // Agent_rate
                    if (pFileModel.Rate > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Rate", pFileModel.Rate);
                        //if (!string.IsNullOrEmpty(pFileModel.PayoutCurrency) )
                        //{
                        //    decimal MinRateLimit = 0m;
                        //    decimal MiaxRateLimit = 0m;
                        //    DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + pFileModel.PayoutCurrency + "'.<br>");
                        //    if (objCurrency != null)
                        //    {
                        //        if (objCurrency.Rows.Count > 0)
                        //        {
                        //            DataRow drCurrency = objCurrency.Rows[0];
                        //            MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                        //            MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                        //        }
                        //        if (pFileModel.Rate < MinRateLimit || pFileModel.Rate > MiaxRateLimit)
                        //        {
                        //            IsRowValid = false;
                        //            err_msgs += "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>";
                        //        }
                        //    }
                        //}
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Rate is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Rate", "");
                    }
                    // PayOut_CCY
                    if (!string.IsNullOrEmpty(pFileModel.PayoutCurrency))
                    {
                        updateCommand.Parameters.AddWithValue("@PayoutCurrency", pFileModel.PayoutCurrency);
                        string PayoutCurrency = pFileModel.PayoutCurrency;
                        if (Common.toString(PayoutCurrency).Trim() != "")
                        {
                            if (Common.isCurrencyAvailable(PayoutCurrency))
                            {
                                string accountcode = Common.getHoldAccountByCurrency(PayoutCurrency, "1");
                                if (accountcode == "")
                                {
                                    AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(PayoutCurrency, "1");
                                    if (!response.isAdded)
                                    {
                                        IsRowValid = false;
                                        err_msgs += PayoutCurrency + " Unpaid Account is not available and unable to create it. Please check currency settings.<br>";
                                    }
                                }
                            }
                            else
                            {
                                IsRowValid = false;
                                err_msgs += PayoutCurrency + " is not available or is inactive. Please check currency settings.<br>";
                            }
                        }
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Payout currency is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@PayoutCurrency", "");
                    }

                    // Amount
                    if (pFileModel.FCAmount > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@FCAmount", pFileModel.FCAmount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "F/C amount is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@FCAmount", "0");
                    }
                    // Payin_amount
                    if (pFileModel.Payinamount >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Payinamount", pFileModel.Payinamount);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Payin amount is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Payinamount", "0");
                    }
                    /// Admin Charges
                    if (pFileModel.AdminCharges >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@AdminCharges", pFileModel.AdminCharges);
                        ////if (pFileModel.AdminCharges > 0 && Common.toString(pFileModel.PayoutCurrency) != "")
                        ////{
                        ////    string accountcode = Common.getHoldAccountByCurrency(pFileModel.PayoutCurrency, "2");
                        ////    if (accountcode == "")
                        ////    {
                        ////        AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(pFileModel.PayoutCurrency, "2");
                        ////        if (!response.isAdded)
                        ////        {
                        ////            IsRowValid = false;
                        ////            err_msgs += pFileModel.PayoutCurrency + " Un-earned admin charges account is not available and unable to create it. Please check currency settings.<br>";
                        ////        }
                        ////    }
                        ////}
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@AdminCharges", "0");
                    }

                    ////     Agent Charges 
                    if (pFileModel.AgentCharges >= 0)
                    {
                        updateCommand.Parameters.AddWithValue("@AgentCharges", pFileModel.AgentCharges);
                        ////if (pFileModel.AgentCharges > 0 && Common.toString(pFileModel.PayoutCurrency) != "")
                        ////{
                        ////    string accountcode = Common.getHoldAccountByCurrency(pFileModel.PayoutCurrency, "3");
                        ////    if (accountcode == "")
                        ////    {
                        ////        AddAccountCodeResponse response = Common.AutoCreateAccountByCurrency(pFileModel.PayoutCurrency, "3");
                        ////        if (!response.isAdded)
                        ////        {
                        ////            IsRowValid = false;
                        ////            err_msgs += pFileModel.PayoutCurrency + " Un-earned agent charges account is not available and unable to create it. Please check currency settings.<br>";
                        ////        }
                        ////    }
                        ////}
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@AgentCharges", "0");
                    }
                    /// Phone
                    if (!string.IsNullOrEmpty(pFileModel.Phone))
                    {
                        updateCommand.Parameters.AddWithValue("@Phone", pFileModel.Phone);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Phone is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Phone", "");
                    }
                    /// FatherName
                    if (!string.IsNullOrEmpty(pFileModel.FatherName))
                    {
                        updateCommand.Parameters.AddWithValue("@FatherName", pFileModel.FatherName);
                    }
                    else
                    {
                        //IsRowValid = false;
                        //err_msgs += "FatherName is not available in excel sheet;";
                        updateCommand.Parameters.AddWithValue("@FatherName", "");
                    }
                    /// Address
                    if (!string.IsNullOrEmpty(pFileModel.Address))
                    {
                        updateCommand.Parameters.AddWithValue("@Address", pFileModel.Address);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Address is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Address", "");
                    }
                    /// City
                    if (!string.IsNullOrEmpty(pFileModel.City))
                    {
                        updateCommand.Parameters.AddWithValue("@City", pFileModel.City);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "City is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@City", "");
                    }
                    /// Code
                    if (!string.IsNullOrEmpty(pFileModel.Code))
                    {
                        updateCommand.Parameters.AddWithValue("@Code", pFileModel.Code);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Code is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Code", "");
                    }
                    /// PaymentNo
                    if (!string.IsNullOrEmpty(pFileModel.PaymentNo))
                    {
                        updateCommand.Parameters.AddWithValue("@PaymentNo", pFileModel.PaymentNo);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "PaymentNo is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@PaymentNo", "");
                    }
                    /// BenePaymentMethod
                    if (!string.IsNullOrEmpty(pFileModel.BenePaymentMethod))
                    {
                        updateCommand.Parameters.AddWithValue("@BenePaymentMethod", pFileModel.BenePaymentMethod);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Bene Payment Method is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@BenePaymentMethod", "");
                    }
                    /// SendingCountry
                    if (!string.IsNullOrEmpty(pFileModel.SendingCountry))
                    {
                        updateCommand.Parameters.AddWithValue("@SendingCountry", pFileModel.SendingCountry);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Sending Country is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@SendingCountry", "");
                    }
                    /// ReceivingCountry
                    if (!string.IsNullOrEmpty(pFileModel.ReceivingCountry))
                    {
                        updateCommand.Parameters.AddWithValue("@ReceivingCountry", pFileModel.ReceivingCountry);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Receiving Country is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@ReceivingCountry", "");
                    }
                    /// Tr_No
                    if (pFileModel.TRANID > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@TRANID", pFileModel.TRANID);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Transaction id is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@TRANID", "");
                    }
                    /// Customer_id
                    if (pFileModel.CustomerId > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerId", pFileModel.CustomerId);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Customer id is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@CustomerId", "");
                    }
                    /// Customer_full_name
                    if (!string.IsNullOrEmpty(pFileModel.CustomerName))
                    {
                        updateCommand.Parameters.AddWithValue("@CustomerName", pFileModel.CustomerName);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Customer name is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@CustomerName", "");
                    }
                    /// Beneficiary_full_name
                    if (!string.IsNullOrEmpty(pFileModel.Recipient))
                    {
                        updateCommand.Parameters.AddWithValue("@Recipient", pFileModel.Recipient);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Recipient is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Recipient", "");
                    }

                    if (!string.IsNullOrEmpty(pFileModel.Status))
                    {
                        updateCommand.Parameters.AddWithValue("@Status", pFileModel.Status);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Status is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Status", "");
                    }
                    if (pFileModel.Charges > 0)
                    {
                        updateCommand.Parameters.AddWithValue("@Charges", pFileModel.Charges);
                    }
                    else
                    {
                        IsRowValid = false;
                        err_msgs += "Charges is not available in excel sheet.<br>";
                        updateCommand.Parameters.AddWithValue("@Charges", "");
                    }
                    if (IsRowValid)
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", true);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                        updateCommand.Parameters.AddWithValue("@RowErrMessage", err_msgs);
                    }
                    else
                    {
                        updateCommand.Parameters.AddWithValue("@isRowValid", false);
                        updateCommand.Parameters.AddWithValue("@RowErrMessage", err_msgs);
                        updateCommand.Parameters.AddWithValue("@Id", pFileModel.Id);
                    }
                    updateCommand.CommandText = UpdateSQL;
                    updateCommand.ExecuteNonQuery();
                    updateCommand.Parameters.Clear();

                }
                catch (Exception ex)
                {

                }
                finally
                {
                    con.Close();
                }
            }
            pFileModel.isRowValid = IsRowValid;
            pFileModel.RowErrMessage = err_msgs;
            return pFileModel;
        }
        public ActionResult PostCancelledData(ImportCancelledFileViewModel pPostingModel)
        {
            ErrorMessages objErrorMessage = null;
            //// string myConnectionString = BaseModel.getConnString();
            string myConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings[Common.toString(SysPrefs.DefaultConnectionString)].ToString();
            ApplicationUser profile = ApplicationUser.GetUserProfile();
            DataSet DS_lstGLEntriesTemp = DBUtils.GetDataSet("select * from FileEntriesTemp where cpUserRefNo='" + pPostingModel.cpUserRefNo + "'");
            DataTable lstGLEntriesTemp = DS_lstGLEntriesTemp.Tables[0];
            if (lstGLEntriesTemp != null)
            {

                if (lstGLEntriesTemp.Rows.Count > 0)
                {
                    bool isMainValid = true;
                    foreach (DataRow objGLEntriesTemp in lstGLEntriesTemp.Rows)
                    {
                        if (Common.toBool(objGLEntriesTemp["isRowValid"]) == false)
                        {
                            isMainValid = false;
                        }
                    }
                    if (isMainValid)
                    {
                        objErrorMessage = PostingUtils.SaveCancelledTransactionBridge("22CANCELENTRY", lstGLEntriesTemp, Profile.Id.ToString()); //Fix type
                        string postingErr = "";
                        if (objErrorMessage.IsFailed)
                        {
                            for (int i = 0; i < objErrorMessage.ErrorMessage.Count; i++)
                            {
                                postingErr = postingErr + "<br>" + objErrorMessage.ErrorMessage[i].Message;
                            }
                            pPostingModel.ErrMessage = Common.GetAlertMessage(1, postingErr);
                        }
                        if (objErrorMessage.ErrorMessage.Count > 0)
                        {
                            pPostingModel.ErrMessage = objErrorMessage.ErrorMessage[0].Message.ToString();
                        }
                        else
                        {
                            pPostingModel.ErrMessage = Common.GetAlertMessage(0, "Hold Entry from Excel is successfully added <br> <br> " + objErrorMessage.ErrorMessage[0].Message + "  <br> <br> <a href ='/Inquiries/GLInquiry'>Click here</a>");
                        }
                    }
                    else
                    {
                        pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Please check your data and try again");
                    }
                }
            }
            else
            {
                pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Data not found");
            }

            return View("~/Views/ImportExcel/MessageCancelled.cshtml", pPostingModel);
        }
        #endregion
    }
}