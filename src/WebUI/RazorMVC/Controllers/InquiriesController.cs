using ClosedXML.Excel;
using Dapper;
using DHAAccounts.Models;
using Kendo.Mvc;
using Kendo.Mvc.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace Accounting.Controllers
{
    public class InquiriesController : Controller
    {
        // private iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
        ApplicationUser Profile = ApplicationUser.GetUserProfile();
        private AppDbContext _dbcontext = new AppDbContext();
        //// private AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());

        string ExportHtml = string.Empty;
        // GET: Inquiries
        public ActionResult Index()
        {
            return View();
        }

        #region GL Inquiry V2
        public ActionResult GLInquiry()
        {
            InquiriesViewModel pInquiriesViewModel = new InquiriesViewModel();
            pInquiriesViewModel.FromDate = Common.toDateTime(SysPrefs.PostingDate).AddDays(-30);
            pInquiriesViewModel.ToDate = Common.toDateTime(SysPrefs.PostingDate);
            pInquiriesViewModel.table = "";
            pInquiriesViewModel.haveData = false;
            pInquiriesViewModel.ExcelCurrencyType = "0";
            pInquiriesViewModel.PdfCurrencyType = "0";
            pInquiriesViewModel.LocalCurrencyISOCode = Common.GetSysSettings("DefaultCurrency");
            return View("~/Views/Inquiries/GLInquiryV2.cshtml", pInquiriesViewModel);
        }
        public JsonResult GetGLAccountsV2([DataSourceRequest] DataSourceRequest request)
        {
            var bankAccount = PostingUtils.GetGeneralAccountsWithCurrency();
            return Json(bankAccount, JsonRequestBehavior.AllowGet);
        }

        public ActionResult CheckList_Read([DataSourceRequest] DataSourceRequest request)
        {

            #region Get Filter Values
            string ParamAccountCode = "";
            string ParamFromDate = "";
            string ParamToDate = "";
            string ParamMinAmount = "";
            string ParamMaxAmount = "";

            if (request.Filters != null && request.Filters.Any())
            {
                foreach (var Filter in request.Filters)
                {
                    var descriptor = Filter as FilterDescriptor;

                    if (descriptor != null && descriptor.Member != "")
                    {
                        if (descriptor.Member == "AccountCode")
                        {
                            ParamAccountCode = Common.toString(descriptor.Value);
                        }
                        if (descriptor.Member == "AddedDate")
                        {
                            ParamFromDate = Common.toString(descriptor.Value);
                        }
                        if (descriptor.Member == "AddedDate")
                        {
                            ParamToDate = Common.toString(descriptor.Value);
                        }
                        if (descriptor.Member == "MinAmount")
                        {
                            ParamMinAmount = Common.toString(descriptor.Value);
                        }
                        if (descriptor.Member == "MaxAmount")
                        {
                            ParamMaxAmount = Common.toString(descriptor.Value);
                        }

                    }
                    else if (Filter is CompositeFilterDescriptor)
                    {
                        var filterDescriptors = ((CompositeFilterDescriptor)Filter).FilterDescriptors;
                        foreach (var filterDescriptor in filterDescriptors)
                        {
                            var descriptor2 = filterDescriptor as FilterDescriptor;

                            if (descriptor2 != null && descriptor2.Member != "")
                            {
                                if (descriptor2.Member == "AccountCode")
                                {
                                    ParamAccountCode = Common.toString(descriptor2.Value);
                                }
                                //if (descriptor.Member == "AddedDate")
                                //{
                                //    ParamFromDate = Common.toString(descriptor.Value);
                                //}
                                //if (descriptor.Member == "AddedDate")
                                //{
                                //    ParamToDate = Common.toString(descriptor.Value);
                                //}
                                //if (descriptor.Member == "MinAmount")
                                //{
                                //    ParamMinAmount = Common.toString(descriptor.Value);
                                //}
                                //if (descriptor.Member == "MaxAmount")
                                //{
                                //    ParamMaxAmount = Common.toString(descriptor.Value);
                                //}
                            }
                            else if (filterDescriptor is CompositeFilterDescriptor)
                            {
                                var filterDescriptors1 = ((CompositeFilterDescriptor)filterDescriptor).FilterDescriptors;
                                {
                                    foreach (var filterDescriptor1 in filterDescriptors1)
                                    {
                                        var descriptor3 = filterDescriptor1 as FilterDescriptor;

                                        if (descriptor3 != null && descriptor3.Member != "")
                                        {
                                            if (descriptor3.Member == "AccountCode")
                                            {
                                                ParamAccountCode = Common.toString(descriptor3.Value);
                                            }
                                            if (descriptor3.Member == "AddedDate")
                                            {
                                                if (Common.toDateTime(ParamFromDate) == null || Common.toString(ParamFromDate) == "")
                                                {
                                                    ParamFromDate = Common.toString(descriptor3.Value);
                                                }
                                            }
                                            if (descriptor3.Member == "AddedDate")
                                            {
                                                ParamToDate = Common.toString(descriptor3.Value);
                                            }
                                            if (descriptor3.Member == "MinAmount")
                                            {
                                                ParamMinAmount = Common.toString(descriptor3.Value);
                                            }
                                            if (descriptor3.Member == "MaxAmount")
                                            {
                                                ParamMaxAmount = Common.toString(descriptor3.Value);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion
            if (!string.IsNullOrEmpty(Common.toString(ParamAccountCode)))
            {
                int CurrentPage = 0; int CurrentPageSize = 1;
                if (request.PageSize > 0)
                {
                    CurrentPageSize = request.PageSize;
                }
                if (request.PageSize > 0 && request.Page > 0)
                {
                    CurrentPage = request.Page - 1;
                }

                AccountBalance AcBalances = PostingUtils.GetOpeningBalance(ParamAccountCode, Common.toDateTime(ParamFromDate));
                decimal LocalCurrencyBalance = Common.toDecimal(AcBalances.LocalCurrencyBalance);
                decimal ForeignCurrencyBalance = Common.toDecimal(AcBalances.ForeignCurrencyBalance);

                if (CurrentPage > 0)
                {
                    string selectQueryLast = @"SELECT top 1  ";
                    selectQueryLast += @"SUM (LocalCurrencyAmount) OVER (ORDER BY RowNum) + (" + LocalCurrencyBalance + ") AS LocalCurrencyBalance";
                    selectQueryLast += @",SUM (ForeignCurrencyAmount) OVER (ORDER BY RowNum) + (" + ForeignCurrencyBalance + ") AS ForeignCurrencyBalance";
                    selectQueryLast += @" FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum";
                    selectQueryLast += @",ForeignCurrencyAmount, LocalCurrencyAmount ";
                    selectQueryLast += @" FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID  /**where**/ ) AS RowConstrainedResult";
                    selectQueryLast += @" WHERE   RowNum <= (@PageIndex) * @PageSize ORDER BY RowNum desc";

                    SqlBuilder builderLast = new SqlBuilder();
                    var selectorLast = builderLast.AddTemplate(selectQueryLast, new { PageIndex = CurrentPage, PageSize = CurrentPageSize });

                    if (request.Filters != null && request.Filters.Any())
                    {
                        builderLast = Common.ApplyFilters(builderLast, request.Filters);
                    }
                    if (request.Sorts != null && request.Sorts.Any())
                    {
                        builderLast = Common.ApplySorting(builderLast, request.Sorts);
                    }
                    else
                    {
                        builderLast.OrderBy("GLTransactions.AddedDate");
                        builderLast.OrderBy("GLReferences.VoucherNumber");
                    }

                    var rowsLast = _dbcontext.Query<InquiriesViewModel>(selectorLast.RawSql, selectorLast.Parameters);
                    if (rowsLast != null)
                    {
                        foreach (InquiriesViewModel modelLast in rowsLast)
                        {
                            LocalCurrencyBalance = Common.toDecimal(modelLast.LocalCurrencyBalance);
                            ForeignCurrencyBalance = Common.toDecimal(modelLast.ForeignCurrencyBalance);
                        }
                    }
                }

                const string countQuery = @"SELECT COUNT(1) FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID  /**where**/";
                string selectQuery = @"SELECT  * ";
                selectQuery += @",SUM (LocalCurrencyAmount) OVER (ORDER BY RowNum) + (" + LocalCurrencyBalance + ") AS LocalCurrencyBalance";
                selectQuery += @",SUM (ForeignCurrencyAmount) OVER (ORDER BY RowNum) + (" + ForeignCurrencyBalance + ") AS ForeignCurrencyBalance";
                selectQuery += @" FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum";
                selectQuery += @",GLTransactions.GLReferenceID, GLTransactions.TransactionID ,GLTransactions.TransactionNumber";
                selectQuery += @",GLTransactions.AddedDate,GLTransactions.AccountCode ,GLTransactions.Memo, GLReferences.VoucherNumber, GLTransactions.TransactionNumber As TransNo";
                selectQuery += @",2 as SortOrder, ForeignCurrencyISOCode, ForeignCurrencyAmount, LocalCurrencyAmount ";
                //selectQuery += @",SUM (LocalCurrencyAmount) OVER (ORDER BY GLTransactions.GLReferenceID) + (" + LocalCurrencyBalance + ") AS LocalCurrencyBalance";
                //selectQuery += @",SUM (ForeignCurrencyAmount) OVER (ORDER BY GLTransactions.GLReferenceID) + (" + ForeignCurrencyBalance + ") AS ForeignCurrencyBalance";
                selectQuery += @",(CASE WHEN GLTransactions.ForeignCurrencyAmount < 0 THEN (GLTransactions.ForeignCurrencyAmount * (-1)) END) AS CreditForeignCurrencyAmount";
                selectQuery += @",(CASE WHEN GLTransactions.ForeignCurrencyAmount >= 0 THEN (GLTransactions.ForeignCurrencyAmount) END) AS DebitForeignCurrencyAmount";
                selectQuery += @",(CASE WHEN GLTransactions.LocalCurrencyAmount < 0 THEN (GLTransactions.LocalCurrencyAmount * (-1)) END) AS CreditLocalCurrencyAmount";
                selectQuery += @",(CASE WHEN GLTransactions.LocalCurrencyAmount >= 0 THEN (GLTransactions.LocalCurrencyAmount) END) AS DebitLocalCurrencyAmount";
                selectQuery += @" FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID  /**where**/ ) AS RowConstrainedResult";
                selectQuery += @" WHERE   RowNum >= (@PageIndex * @PageSize + 1 ) AND RowNum <= (@PageIndex + 1) * @PageSize ORDER BY RowNum";

                SqlBuilder builder = new SqlBuilder();
                var count = builder.AddTemplate(countQuery);

                var selector = builder.AddTemplate(selectQuery, new { PageIndex = CurrentPage, PageSize = CurrentPageSize });

                if (request.Filters != null && request.Filters.Any())
                {
                    builder = Common.ApplyFilters(builder, request.Filters);
                }
                if (request.Sorts != null && request.Sorts.Any())
                {
                    builder = Common.ApplySorting(builder, request.Sorts);
                }
                else
                {
                    builder.OrderBy("GLTransactions.AddedDate");
                    builder.OrderBy("GLReferences.VoucherNumber");
                }

                var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);
                var rows = _dbcontext.Query<InquiriesViewModel>(selector.RawSql, selector.Parameters);

                Session["PageTotal"] = 0;

                foreach (InquiriesViewModel model in rows)
                {
                    model.GLReferenceID = Security.EncryptQueryString(model.GLReferenceID).ToString();
                    //string AccountName = Common.toString(DBUtils.executeSqlGetSingle("select AccountName from GLChartOfAccounts where AccountCode='"+model.AccountCode+"'"));
                    //if (model.TransactionNumber.Trim()== "Inc-Revaluation")
                    //{
                    //    model.DebitLocalCurrencyAmount = model.CreditLocalCurrencyAmount;
                    //    model.CreditLocalCurrencyAmount =0;
                    //}
                }

                #region Opening Balance  
                List<InquiriesViewModel> openingBalance = new List<InquiriesViewModel>();
                if (CurrentPage == 0)
                {
                    if (!string.IsNullOrEmpty(Common.toString(ParamAccountCode)))
                    {
                        //AccountBalance AcBalances = PostingUtils.GetOpeningBalance(ParamAccountCode, Common.toDateTime(ParamFromDate));
                        //decimal LocalCurrencyBalance = Common.toDecimal(AcBalances.LocalCurrencyBalance);
                        //decimal ForeignCurrencyBalance = Common.toDecimal(AcBalances.ForeignCurrencyBalance);
                        if (!string.IsNullOrEmpty(Common.toString(ParamAccountCode)))
                        {
                            openingBalance.Add(new InquiriesViewModel
                            {
                                VoucherNumber = "Opening Balance - " + Common.toDateTime(ParamFromDate).ToShortDateString(),
                                LocalCurrencyBalance = LocalCurrencyBalance,
                                ForeignCurrencyBalance = ForeignCurrencyBalance,
                                SortOrder = 1,
                            });
                        }
                    }
                }
                #endregion

                #region Closing Balance
                decimal LastPage = 0;
                decimal total = Common.toDecimal(totalCount);
                decimal rem = total / 50 % 1;
                if (rem == 0)
                {
                    LastPage = Math.Floor(total / 50);
                    LastPage = LastPage - 1;
                }
                else
                {
                    LastPage = Math.Floor(total / 50);
                }
                List<InquiriesViewModel> closingBalance = new List<InquiriesViewModel>();
                if (CurrentPage == LastPage)
                {
                    decimal ClosingLocalCurrencyBalance = 0;
                    decimal ClosingForeignCurrencyBalance = 0;
                    foreach (InquiriesViewModel model in rows)
                    {
                        ClosingLocalCurrencyBalance = model.LocalCurrencyBalance;
                        ClosingForeignCurrencyBalance = model.ForeignCurrencyBalance;
                    }
                    if (!string.IsNullOrEmpty(Common.toString(ParamAccountCode)))
                    {
                        closingBalance.Add(new InquiriesViewModel
                        {
                            VoucherNumber = "Closing Balance - " + Common.toDateTime(ParamToDate).ToShortDateString(),
                            LocalCurrencyBalance = ClosingLocalCurrencyBalance,
                            ForeignCurrencyBalance = ClosingForeignCurrencyBalance,
                            SortOrder = 3,
                        });
                    }
                }
                #endregion

                var newList = rows.Concat(closingBalance);
                newList = openingBalance.Concat(newList);

                //foreach(var item in newList)
                //{
                //    item.AccountCodeWithMemo = $"{item.AccountCode} - {item.Memo}";
                //}
                var result = new DataSourceResult()
                {
                    Data = newList,
                    Total = totalCount
                };
                return Json(result);
            }
            else
            {
                var result = new DataSourceResult()
                {
                    Data = "",
                    Total = 0
                };
                return Json(result);
            }
        }

        #endregion

        #region GL Trans View
        public ActionResult GLTransView(string id)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            TransViewModel pGLTransViewModel = new TransViewModel();
            if (!string.IsNullOrEmpty(id))
            {
                pGLTransViewModel.cpRefID = Security.DecryptQueryString(id);
                pGLTransViewModel.HtmlTable1 = BuildDynamicTable1(Common.toInt(pGLTransViewModel.cpRefID));
                string Records = "Select VoucherNumber, TypeiCode, AuthorizedBy, AuthorizedDate, DatePosted, PostedBy from GLReferences where GLReferenceID =" + pGLTransViewModel.cpRefID + "";
                DataTable transTable = DBUtils.GetDataTable(Records);
                if (transTable != null && transTable.Rows.Count > 0)
                {
                    foreach (DataRow RecordTable in transTable.Rows)
                    {
                        pGLTransViewModel.AuthorizedBy = Common.GetUserFullName(Common.toString(RecordTable["AuthorizedBy"]));
                        pGLTransViewModel.AuthorizedDate = Common.toDateTime(RecordTable["AuthorizedDate"]).ToShortDateString();
                        if (pGLTransViewModel.AuthorizedDate == "01/01/1990")
                        {
                            pGLTransViewModel.AuthorizedDate = "";
                        }
                        pGLTransViewModel.DatePosted = Common.toDateTime(RecordTable["DatePosted"]).ToShortDateString();
                        pGLTransViewModel.PostedBy = Common.GetUserFullName(Common.toString(RecordTable["PostedBy"]));
                        pGLTransViewModel.VoucherNumber = Common.toString(RecordTable["VoucherNumber"]);
                        pGLTransViewModel.TypeiCode = Common.toString(RecordTable["TypeiCode"]);
                        pGLTransViewModel.VoucherTitle = GetVoucherTitle(pGLTransViewModel.TypeiCode);
                    }
                }
            }
            string getFileName = "Select FileName from GLReferencesAttachment where GLReferenceID='" + pGLTransViewModel.cpRefID + "'";
            DataTable dtFileName = DBUtils.GetDataTable(getFileName);
            if (dtFileName != null && dtFileName.Rows.Count > 0)
            {
                foreach (DataRow drFileName in dtFileName.Rows)
                {
                    // pGLTransViewModel.Attachments += Common.toString(drFileName["FileName"]) + ", ";
                    //pGLTransViewModel.Attachments += "<a href ='/Uploads/Vouchers/" + drFileName["FileName"].ToString() + "', new { target = '_blank' }>" + Common.toString(drFileName["FileName"]) + "</a>" + " ,";
                    pGLTransViewModel.Attachments += "<a href ='/Uploads/" + SysPrefs.SubmissionFolder + "/" + drFileName["FileName"].ToString() + "', new { target = '_blank' }>" + Common.toString(drFileName["FileName"]) + "</a>" + " ,";
                }
            }
            return View("~/Views/Inquiries/GLTransView.cshtml", pGLTransViewModel);
        }

        private string GetVoucherTitle(string typeiCode)
        {
            if (typeiCode.ToUpper() == "22CASHRECPT") return "Cash Recpt Entry Voucher";
            if (typeiCode.ToUpper() == "22BANKRECPT") return "Bank Recpt Entry Voucher";
            if (typeiCode.ToUpper() == "22BANKPMT") return "Bank Entry Voucher";
            if (typeiCode.ToUpper() == "22CASHPMT") return "Cash Entry Voucher";
            if (typeiCode.ToUpper() == "22JOURNAL") return "Journal Entry Voucher";
            return "General Voucher";
        }

        //public void VoucherToPrint(TranslationGainLossModel pModel)
        //{
        //    TransViewModel pGLTransViewModel = new TransViewModel();
        //    if (!string.IsNullOrEmpty(pModel.cpRefID))
        //    {
        //        pGLTransViewModel.cpRefID = pModel.cpRefID;
        //        pGLTransViewModel.HtmlTable1 = BuildDynamicTable1(Common.toInt(pModel.cpRefID));
        //        string Records = "Select VoucherNumber, AuthorizedBy, AuthorizedDate, DatePosted, PostedBy from GLReferences where GLReferenceID=" + pGLTransViewModel.cpRefID + "";
        //        DataTable transTable = DBUtils.GetDataTable(Records);
        //        if (transTable != null && transTable.Rows.Count > 0)
        //        {
        //            foreach (DataRow RecordTable in transTable.Rows)
        //            {
        //                pGLTransViewModel.AuthorizedBy = Common.GetUserFullName(Common.toString(RecordTable["AuthorizedBy"]));
        //                pGLTransViewModel.AuthorizedDate = Common.toDateTime(RecordTable["AuthorizedDate"]).ToShortDateString();
        //                if (pGLTransViewModel.AuthorizedDate == "01/01/1990")
        //                {
        //                    pGLTransViewModel.AuthorizedDate = "";
        //                }
        //                pGLTransViewModel.DatePosted = Common.toDateTime(RecordTable["DatePosted"]).ToShortDateString();
        //                pGLTransViewModel.PostedBy = Common.GetUserFullName(Common.toString(RecordTable["PostedBy"]));
        //                pGLTransViewModel.VoucherNumber = Common.toString(RecordTable["VoucherNumber"]);
        //            }
        //        }
        //        string getFileName = "Select FileName from GLReferencesAttachment where GLReferenceID='" + pGLTransViewModel.cpRefID + "'";
        //        DataTable dtFileName = DBUtils.GetDataTable(getFileName);
        //        if (dtFileName != null && dtFileName.Rows.Count > 0)
        //        {
        //            foreach (DataRow drFileName in dtFileName.Rows)
        //            {
        //                pGLTransViewModel.Attachments += "<a href ='/Uploads/" + SysPrefs.SubmissionFolder + "/" + drFileName["FileName"].ToString() + "', new { target = '_blank' }>" + Common.toString(drFileName["FileName"]) + "</a>" + " ,";
        //            }
        //        }

        //        string HostName = System.Web.HttpContext.Current.Request.Url.Host;

        //        string htmltbl = "";

        //        htmltbl += "<div class='stm-pnl' style='line-height: 1.5;width: 90%;margin: 2% auto auto auto;'>";
        //        htmltbl += "<div class='container-fluid' style='margin-left: auto;margin-right: auto;padding-left: 15px;padding-right: 15px;';>";

        //        htmltbl += "<div class='row'>";
        //        htmltbl += "<div class='col-md-6' style='float: left;width: 50%;'>";
        //        htmltbl += "<div class='media' style='overflow: hidden;'>";
        //        htmltbl += "<div class='media-body' style='display: table-cell;vertical-align: top;'>";
        //        htmltbl += "<h4 class='media-heading' style='color:darkred;'>" + @SysPrefs.SiteName + "</h4>";
        //        htmltbl += "</div>";
        //        htmltbl += "</div>";
        //        // htmltbl += "<br/><br/>";
        //        //htmltbl += "<div align='center' class='text-center'><strong>Duplicate Voucher</strong></div>";
        //        //htmltbl += "<br/><br/>";
        //        htmltbl += "<h4><b><u>Transfer Entry Voucher</u></b>       <span style='margin-left:200px;'>( Duplicate Voucher )</span></h4>";
        //        htmltbl += "</div>";
        //        htmltbl += "<div class='col-md-6' style='float: left;width: 50%;margin-top:50px;'>";
        //        htmltbl += "<div style='border:#CCC 1px solid; padding:10px;'>";
        //        htmltbl += "<table align='center' style='border-collapse: collapse;background-color: transparent;'>";
        //        htmltbl += "<tr>";
        //        htmltbl += "<td><h5><b>Voucher No.</b></h5></td>";
        //        htmltbl += "<td></td>";
        //        htmltbl += "<td>";
        //        htmltbl += "<h5><b>" + pGLTransViewModel.VoucherNumber + "</b></h5>";
        //        htmltbl += " </td>";
        //        htmltbl += "</tr>";
        //        htmltbl += "<tr>";
        //        htmltbl += "<td>Print Date:</td>";
        //        htmltbl += "<td></td>";
        //        htmltbl += "<td>" + @DateTime.Now.ToShortDateString() + " </td>";
        //        htmltbl += "</tr>";
        //        htmltbl += "<tr>";
        //        htmltbl += "<td>Entry Date:</td>";
        //        htmltbl += "<td></td>";
        //        htmltbl += "<td>" + pGLTransViewModel.DatePosted + "</td>";
        //        htmltbl += "</tr>";
        //        htmltbl += "<tr>";
        //        htmltbl += "<td>Entry By:</td>";
        //        htmltbl += "<td></td>";
        //        htmltbl += "<td>" + pGLTransViewModel.PostedBy + "</td>";
        //        htmltbl += "</tr>";
        //        htmltbl += "<tr>";
        //        htmltbl += "<td>Authorize Date:</td>";
        //        htmltbl += "<td></td>";
        //        htmltbl += "<td>" + pGLTransViewModel.AuthorizedDate + "</td>";
        //        htmltbl += "</tr>";
        //        htmltbl += "<tr>";
        //        htmltbl += "<td>Authorize By:</td>";
        //        htmltbl += "<td></td>";
        //        htmltbl += "<td>" + pGLTransViewModel.AuthorizedBy + "</td>";
        //        htmltbl += "</tr>";
        //        htmltbl += "</table>";
        //        htmltbl += "</div>";
        //        htmltbl += "</div>";
        //        htmltbl += "</div>";
        //        htmltbl += "<label style='display:inline-block;margin-bottom:.5rem;'></label>";
        //        htmltbl += "<span id='lblDetails'>" + pGLTransViewModel.HtmlTable1 + "</span>";
        //        htmltbl += "<div class='form-group'>";
        //        htmltbl += "<div class='col-md-4'></div>";
        //        htmltbl += "<div class='col-md-4'></div>";
        //        htmltbl += "</div>";
        //        if (Common.toString(pGLTransViewModel.Attachments) != "")
        //        {
        //            htmltbl += "<h4>Attachments</h4>";
        //            htmltbl += pGLTransViewModel.Attachments;
        //        }
        //        htmltbl += "</div>";
        //        htmltbl += "</div>";
        //        ////  string fileName = Common.getRandomDigit(4) + "-" + "Voucher";
        //        string fileName = "Voucher-" + pGLTransViewModel.VoucherNumber + "-" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
        //        string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp");
        //        HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //        htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";

        //        PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(htmltbl, sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //        pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //        Response.Clear();
        //        Response.ClearContent();
        //        Response.ClearHeaders();
        //        Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".pdf");
        //        Response.ContentType = "application/pdf";
        //        Response.Flush();
        //        Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".pdf"));

        //        Response.End();
        //    }
        //}

        protected static string BuildDynamicTable1(int cpReferenceID)
        {
            string table = "";
            decimal cpTotalDebit = 0m;
            decimal cpTotalCredit = 0m;
            table += "<table border='1' width='100%' cellpadding='5' style='border-collapse: collapse;'>";
            table += "<thead>";
            table += " <tr>";
            table += "<th style='background-color:#1c3a70; color:#FFF;'>Title of Account/Description</th>";
            table += "<th style='background-color:#1c3a70; color:#FFF; text-align:center;'>Curr</th>";
            table += "<th style='background-color:#1c3a70; color:#FFF; text-align:right;'>FC Amount</th>";
            table += "<th style='background-color:#1c3a70; color:#FFF; text-align:right;'>Debit</th>";
            table += "<th style='background-color:#1c3a70; color:#FFF; text-align:right;'>Credit</th>";
            table += "</tr>";
            table += "<thead>";
            table += "<tbody>";
            DataTable objGLTransView = DBUtils.GetDataTable("select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate from GLTransactions where GLReferenceID='" + cpReferenceID.ToString() + "' order by LocalCurrencyAmount Desc");
            if (objGLTransView != null && objGLTransView.Rows.Count > 0)
            {
                int iRowNumber = 0;
                foreach (DataRow MyRow in objGLTransView.Rows)
                {
                    DataTable objGLReference = DBUtils.GetDataTable("select GLReferenceID, TypeiCode, ReferenceNo, Memo From GLReferences Where GLReferenceID='" + cpReferenceID.ToString() + "'");
                    if (objGLReference != null)
                    {
                        DataRow RowGLReference = objGLReference.Rows[0];
                        string FC = "0.00";
                        if (!string.IsNullOrEmpty(MyRow["ForeignCurrencyAmount"].ToString()))
                        {
                            FC = MyRow["ForeignCurrencyAmount"].ToString();
                        }
                        decimal FcAmount = Math.Abs(Convert.ToDecimal(FC));
                        table += "<tr><td valign='top'><p><strong>" + MyRow["AccountCode"].ToString() + "-" + MyRow["AccountName"].ToString() + "</strong></p>" + MyRow["Memo"].ToString() + "<br/>" + MyRow["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + MyRow["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + RowGLReference["ReferenceNo"].ToString() + "</td>";
                        table += "<td valign='top' align='center' style='border-left:transparent 1px solid;'>" + MyRow["ForeignCurrencyISOCode"].ToString() + "</td>";
                        table += "<td valign='top' align='right'>" + Math.Round(FcAmount, 2) + "</td>";
                        string LC = "0.00";
                        if (!string.IsNullOrEmpty(MyRow["LocalCurrencyAmount"].ToString()))
                        {
                            LC = MyRow["LocalCurrencyAmount"].ToString();
                        }
                        decimal LcAmount = Common.toDecimal(LC);
                        LcAmount = Math.Round(LcAmount, 4);
                        LcAmount = DisplayUtils.getRoundDigit(LC);
                        //if (Common.toString(MyRow["AccountName"]).Trim() == "Revaluation Gain & Loss")
                        //{
                        //    if (LcAmount >= 0)
                        //    {
                        //        LcAmount = -1 * LcAmount;
                        //    }
                        //}
                        if (LcAmount < 0)
                        {
                            table += "<td>&nbsp;</td>";
                            table += "<td valign='top' align='right'>" + Math.Round(-1 * LcAmount, 2) + "</td>";
                            cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                        }
                        else
                        {
                            table += "<td valign='top' align='right'>" + Math.Round(LcAmount, 2) + "</td>";
                            table += "<td>&nbsp;</td>";
                            cpTotalDebit = cpTotalDebit + LcAmount;
                        }
                        table += "</tr>";
                        table += "</tbody>";
                    }
                    iRowNumber++;
                }
            }
            table += "<tfoot>";
            table += "<tr>";
            table += "<td colspan='2' style='background-color:#1c3a70; color:#FFF;'>Total:</td>";
            table += "<td style='background-color:#1c3a70; color:#FFF;'></td>";
            table += "<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + Math.Round(cpTotalDebit, 2) + "</td>";
            table += "<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + Math.Round(cpTotalCredit, 2) + "</td>";
            //table += "<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + DisplayUtils.GetSystemAmountFormat(cpTotalCredit) + "</td>";
            table += "</tr>";
            table += "</tfoot>";
            table += "</table>";
            return table;
        }
        #endregion

        #region GL Inquiry
        //public ActionResult GLInquiry()
        //{
        //    InquiriesViewModel pInquiriesViewModel = new InquiriesViewModel();
        //    pInquiriesViewModel.FromDate = Common.toDateTime(SysPrefs.PostingDate).AddDays(-30);
        //    pInquiriesViewModel.ToDate = Common.toDateTime(SysPrefs.PostingDate);
        //    pInquiriesViewModel.table = "";
        //    pInquiriesViewModel.haveData = false;
        //    pInquiriesViewModel = BuildGLInquiryTable(pInquiriesViewModel);
        //    if (Session["tableHtml"] != null)
        //        Session["tableHtml"] = null;
        //    if (Session["InquiriesViewModel"] != null)
        //        Session["InquiriesViewModel"] = null;
        //    if (Session["objGLInquiry"] != null)
        //        Session["objGLInquiry"] = null;

        //    return View("~/Views/Inquiries/GLInquiry.cshtml", pInquiriesViewModel);
        //}
        //[HttpPost]
        //public ActionResult GLInquiry(InquiriesViewModel pInquiriesViewModel)
        //{
        //    pInquiriesViewModel = BuildGLInquiryTable(pInquiriesViewModel);
        //    return View("~/Views/Inquiries/GLInquiry.cshtml", pInquiriesViewModel);
        //}

        public InquiriesViewModel BuildGLInquiryTable(InquiriesViewModel pInquiriesViewModel)
        {
            pInquiriesViewModel.haveData = false;
            decimal dFCRunningTotal = 0m;
            decimal dLCRunningTotal = 0m;
            StringBuilder strGLInquiryTable = new StringBuilder();
            string myAccountCode = "";
            bool isValid = true;
            if (string.IsNullOrEmpty(pInquiriesViewModel.PdfAccountCode))
            {
                pInquiriesViewModel.ErrMessage = "Please choose account.<br/>";
                isValid = false;
            }
            else
            {
                myAccountCode = pInquiriesViewModel.PdfAccountCode;
            }
            if (pInquiriesViewModel.Dimension == "0" && pInquiriesViewModel.Dimension == "")
            {
                pInquiriesViewModel.ErrMessage = "Please choose dimension.<br/>";
                isValid = false;
            }
            if (pInquiriesViewModel.CurrencyType == null && pInquiriesViewModel.CurrencyType == "")
            {
                pInquiriesViewModel.ErrMessage = "Please choose Currency type.<br/>";
                isValid = false;
            }
            if (isValid)
            {
                //string mySql = "Select GLReferenceID, (Select Description From ConfigItemsData Where ConfigItemsData.ItemsDataCode=GLTransactions.AccountTypeiCode) As TypeiCode, AddedDate, AccountCode + '--' + (Select AccountName From GLChartOfAccounts Where  GLChartOfAccounts.AccountCode=GLTransactions.AccountCode) As AccountName, (Select Title From GLDimensions Where GLDimensions.GLDimensionID=GLTransactions.DimensionId) As Dimension, ForeignCurrencyAmount,LocalCurrencyAmount, (Select Memo From GLReferences Where GLReferences.GLReferenceID=GLTransactions.GLReferenceID) As Memo, Case  When GLPersonTypeiCode IS NULL Then '0' Else GLPersonTypeiCode End As GLPersonTypeiCode, Case  When PersonID IS NULL Then '0' Else PersonID End As PersonID From GLTransactions  Where AccountCode='" + myAccountCode + "'";
                //string mySql = "Select GLReferenceID, AddedDate, ForeignCurrencyAmount, LocalCurrencyAmount, Memo, (Select VoucherNumber From GLReferences Where GLReferences.GLReferenceID=GLTransactions.GLReferenceID) As VoucherNumber, TransactionNumber As TransNo, '' as debit_ForeignCurrencyAmount,'' as debit_LocalCurrencyAmount,'' as credit_ForeignCurrencyAmount,'' as credit_LocalCurrencyAmount ,'' as Bal_ForeignCurrencyAmount, '',  '' as Bal_LocalCurrencyAmount,ForeignCurrencyISOCode From GLTransactions  Where  AccountCode='" + myAccountCode + "'";
                string mySql = "Select (select top 1 agentprefix from CustomerTransactions where PaymentNo=GLTransactions.TransactionNumber) as prefix, GLTransactions.GLReferenceID, GLTransactions.PayoutAmount,GLTransactions.AddedDate, GLTransactions.ForeignCurrencyAmount, GLTransactions.LocalCurrencyAmount, GLTransactions.Memo, GLReferences.VoucherNumber, GLTransactions.TransactionNumber As TransNo, '' as debit_ForeignCurrencyAmount,'' as debit_LocalCurrencyAmount,'' as credit_ForeignCurrencyAmount,'' as credit_LocalCurrencyAmount ,'' as Bal_ForeignCurrencyAmount, '',  '' as Bal_LocalCurrencyAmount,ForeignCurrencyISOCode FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID Where GLReferences.isPosted = 1 and GLTransactions.AccountCode = '" + myAccountCode + "'";

                #region SQL Criteria 
                if (pInquiriesViewModel.PdfFromDate != null)
                {
                    mySql += " AND GLTransactions.AddedDate >='" + pInquiriesViewModel.PdfFromDate.ToString("yyyy-MM-dd") + " 00:00:00'";
                }
                if (pInquiriesViewModel.PdfToDate != null)
                {
                    mySql += " AND GLTransactions.AddedDate <='" + pInquiriesViewModel.PdfToDate.ToString("yyyy-MM-dd") + " 23:59:00'";
                }
                if (pInquiriesViewModel.Dimension != "" && pInquiriesViewModel.Dimension != "0")
                {
                    // mySql += " AND DimensionId = '" + ddlDimension1.SelectedValue.ToString() + "'";
                }

                if (pInquiriesViewModel.PdfCurrencyType != "0")
                {
                    if (pInquiriesViewModel.PdfCurrencyType == "1") // Foreign Currency
                    {
                        if (pInquiriesViewModel.MinAmount > 0)
                        {
                            mySql += " AND GLTransactions.ForeignCurrencyAmount >= '" + pInquiriesViewModel.MinAmount + "'";
                        }
                        if (pInquiriesViewModel.MaxAmount > 0)
                        {
                            mySql += " AND GLTransactions.ForeignCurrencyAmount <= '" + pInquiriesViewModel.MaxAmount + "'";
                        }
                    }
                    else if (pInquiriesViewModel.PdfCurrencyType == "2") // Local Currency
                    {
                        if (pInquiriesViewModel.MinAmount > 0)
                        {
                            mySql += " AND GLTransactions.LocalCurrencyAmount >= '" + pInquiriesViewModel.MinAmount + "'";
                        }
                        if (pInquiriesViewModel.MaxAmount > 0)
                        {
                            mySql += " AND GLTransactions.LocalCurrencyAmount <= '" + pInquiriesViewModel.MaxAmount + "'";
                        }
                    }
                }
                mySql += " Order by GLTransactions.AddedDate, GLReferences.VoucherNumber ";

                #endregion

                DataTable objGLInquiry = DBUtils.GetDataTable(mySql);

                //strGLInquiryTable.Append(GetGLTableHeader(myAccountCode,pInquiriesViewModel));
                ExportHtml = "<div><b> Account Ledger - Both </b></div> <br><br><br> ";
                ExportHtml += "<div style='width: 100%; text-align: right;'><tr> ";
                ExportHtml += "<td>Print Date:	</td><td style='text-align:right'>" + DateTime.Now.ToShortDateString() + "</td>";
                ExportHtml += "<br><br> ";
                ExportHtml += "<td>From Date:	</td><td style='text-align:right'>" + pInquiriesViewModel.FromDate.ToShortDateString() + "</td>";
                ExportHtml += "<br><br> ";
                ExportHtml += "<td>To Date:	</td><td style='text-align:right'>" + pInquiriesViewModel.ToDate.ToShortDateString() + "</td>";
                ExportHtml += "<br><br> ";
                ExportHtml += "</tr></div> ";
                HeaderViewModel headerViewModel = GetGLTableHeader(myAccountCode, pInquiriesViewModel);
                dFCRunningTotal = headerViewModel.FCOpeningBalance;
                dLCRunningTotal = headerViewModel.LCOpeningBalance;
                strGLInquiryTable.Append(headerViewModel.HtmlBody);
                if (objGLInquiry != null && objGLInquiry.Rows.Count > 0)
                {
                    pInquiriesViewModel.haveData = true;
                    foreach (DataRow Item in objGLInquiry.Rows)
                    {
                        strGLInquiryTable.Append("<tr>");
                        strGLInquiryTable.Append("<td><a href =\"javascript: showDetails('/Inquiries/GLTransView?Id=" + Security.EncryptQueryString(Item["GLReferenceID"].ToString()) + "','Voucher View')\")>" + Item["VoucherNumber"].ToString() + "</a></td>");
                        strGLInquiryTable.Append("<td>" + Common.toDateTime(Item["AddedDate"]).ToShortDateString() + "</td>");
                        strGLInquiryTable.Append("<td>" + Item["Memo"].ToString() + "</td>");
                        strGLInquiryTable.Append("<td>" + Item["TransNo"].ToString() + "</td>");
                        strGLInquiryTable.Append("<td>" + Math.Round(Common.toDecimal(Item["PayoutAmount"]), 2) + "</td>");

                        ExportHtml += "<tr>";
                        ExportHtml += "<td> " + Item["VoucherNumber"].ToString() + "</td>";
                        ExportHtml += "<td style='text-align:center'>" + Common.toDateTime(Item["AddedDate"]).ToShortDateString() + "</td>";
                        ExportHtml += "<td style='text-align:center'>" + Item["Memo"].ToString() + "</td>";
                        ExportHtml += "<td style='text-align:center'>" + Item["TransNo"].ToString() + "</td>";
                        ExportHtml += "<td style='text-align:center'>" + Math.Round(Common.toDecimal(Item["PayoutAmount"]), 2) + "</td>";
                        #region Both Currency
                        if (pInquiriesViewModel.PdfCurrencyType == "0")
                        {
                            #region Debit
                            if (Common.toDecimal(Item["ForeignCurrencyAmount"]) > 0)
                            {
                                strGLInquiryTable.Append("<td style='text-align:right'><b>" + DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["ForeignCurrencyAmount"])) + "</b></td>");
                                Item["debit_ForeignCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["ForeignCurrencyAmount"]));
                                ExportHtml += "<td style='text-align:right'><b>" + DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["ForeignCurrencyAmount"])) + "</b></td>";
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td>&nbsp;</td>");

                                ExportHtml += "<td>&nbsp;</td>";
                            }
                            if (Common.toDecimal(Item["LocalCurrencyAmount"]) > 0)
                            {
                                strGLInquiryTable.Append("<td style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["LocalCurrencyAmount"])) + "</td>");
                                Item["debit_LocalCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["LocalCurrencyAmount"]));
                                ExportHtml += "<td style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["LocalCurrencyAmount"])) + "</td>";
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td>&nbsp;</td>");
                                ExportHtml += "<td>&nbsp;</td>";
                            }
                            #endregion

                            #region Credit
                            if (Common.toDecimal(Item["ForeignCurrencyAmount"]) < 0)
                            {
                                strGLInquiryTable.Append("<td style='text-align:right' >" + DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["ForeignCurrencyAmount"])) + "</td>");
                                ExportHtml += "<td style='text-align:right' >" + DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["ForeignCurrencyAmount"])) + "</td>";
                                Item["credit_ForeignCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["ForeignCurrencyAmount"]));
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td>&nbsp;</td>");
                                ExportHtml += "<td>&nbsp;</td>";
                            }
                            if (Common.toDecimal(Item["LocalCurrencyAmount"]) < 0)
                            {
                                strGLInquiryTable.Append("<td style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["LocalCurrencyAmount"])) + "</td>");
                                ExportHtml += "<td style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["LocalCurrencyAmount"])) + "</td>";
                                Item["credit_LocalCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["LocalCurrencyAmount"]));
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td>&nbsp;</td>");
                                ExportHtml += "<td>&nbsp;</td>";
                            }
                            #endregion

                            #region Header
                            dFCRunningTotal += Common.toDecimal(Item["ForeignCurrencyAmount"]);
                            dLCRunningTotal += Common.toDecimal(Item["LocalCurrencyAmount"]);

                            strGLInquiryTable.Append("<td style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(dFCRunningTotal) + "</td>");
                            strGLInquiryTable.Append("<td style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(dLCRunningTotal) + "</td>");
                            Item["Bal_ForeignCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(dFCRunningTotal);
                            Item["Bal_LocalCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(dLCRunningTotal);
                            ExportHtml += "<td style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(dFCRunningTotal) + "</td>";
                            ExportHtml += "<td style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(dLCRunningTotal) + "</td>";

                            if (dLCRunningTotal > 0)
                            {
                                strGLInquiryTable.Append("<td>Dr</td>");
                                ExportHtml += "<td>Dr</td>";
                            }
                            else if (dLCRunningTotal < 0)
                            {
                                strGLInquiryTable.Append("<td>Cr</td>");
                                ExportHtml += "<td>Cr</td>";
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td>&nbsp;</td>");
                                ExportHtml += "<td>&nbsp;</td>";
                            }
                            #endregion
                        }
                        #endregion

                        #region Foreign Currency
                        else if (pInquiriesViewModel.PdfCurrencyType == "1")
                        {
                            #region Debit
                            if (Common.toDecimal(Item["ForeignCurrencyAmount"]) > 0)
                            {
                                strGLInquiryTable.Append("<td colspan=2 style='text-align:right'><b>" + DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["ForeignCurrencyAmount"])) + "</b></td>");
                                Item["debit_ForeignCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["ForeignCurrencyAmount"]));
                                ExportHtml += "<td colspan=2 style='text-align:right'><b>" + DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["ForeignCurrencyAmount"])) + "</b></td>";
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td colspan=2>&nbsp;</td>");
                                ExportHtml += "<td colspan=2>&nbsp;</td>";
                            }

                            #endregion

                            #region Credit
                            if (Common.toDecimal(Item["ForeignCurrencyAmount"]) < 0)
                            {
                                strGLInquiryTable.Append("<td colspan=2 style='text-align:right' >" + DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["ForeignCurrencyAmount"])) + "</td>");
                                Item["credit_ForeignCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["ForeignCurrencyAmount"]));
                                ExportHtml += "<td colspan=2 style='text-align:right' >" + DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["ForeignCurrencyAmount"])) + "</td>";
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td colspan=2>&nbsp;</td>");
                                ExportHtml += "<td colspan=2>&nbsp;</td>";
                            }

                            #endregion

                            #region Header
                            dFCRunningTotal += Common.toDecimal(Item["ForeignCurrencyAmount"]);
                            strGLInquiryTable.Append("<td  colspan=2 style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(dFCRunningTotal) + "</td>");
                            Item["Bal_ForeignCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(dFCRunningTotal);
                            ExportHtml += "<td  colspan=2 style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(dFCRunningTotal) + "</td>";

                            if (dFCRunningTotal > 0)
                            {
                                strGLInquiryTable.Append("<td>Dr</td>");
                                ExportHtml += "<td>Dr</td>";
                            }
                            else if (dFCRunningTotal < 0)
                            {
                                strGLInquiryTable.Append("<td>Cr</td>");
                                ExportHtml += "<td>Cr</td>";
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td>&nbsp;</td>");
                                ExportHtml += "<td>&nbsp;</td>";
                            }
                            #endregion
                        }
                        #endregion

                        #region Local Currency
                        else if (pInquiriesViewModel.PdfCurrencyType == "2")
                        {

                            #region Debit
                            if (Common.toDecimal(Item["LocalCurrencyAmount"]) > 0)
                            {
                                strGLInquiryTable.Append("<td colspan=2 style='text-align:right'><b>" + DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["LocalCurrencyAmount"])) + "</b></td>");
                                Item["debit_LocalCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["LocalCurrencyAmount"]));
                                ExportHtml += "<td colspan=2 style='text-align:right'><b>" + DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["LocalCurrencyAmount"])) + "</b></td>";
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td colspan=2>&nbsp;</td>");
                                ExportHtml += "<td colspan=2>&nbsp;</td>";
                            }

                            #endregion

                            #region Credit
                            if (Common.toDecimal(Item["LocalCurrencyAmount"]) < 0)
                            {
                                strGLInquiryTable.Append("<td colspan=2 style='text-align:right' >" + DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["LocalCurrencyAmount"])) + "</td>");
                                ExportHtml += "<td colspan=2 style='text-align:right' >" + DisplayUtils.GetSystemAmountFormat(-1 * Common.toDecimal(Item["LocalCurrencyAmount"])) + "</td>";
                                Item["credit_LocalCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(Common.toDecimal(Item["LocalCurrencyAmount"]));
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td colspan=2>&nbsp;</td>");
                                ExportHtml += "<td colspan=2>&nbsp;</td>";
                            }

                            #endregion

                            #region Header
                            dLCRunningTotal += Common.toDecimal(Item["LocalCurrencyAmount"]);
                            strGLInquiryTable.Append("<td  colspan=2 style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(dLCRunningTotal) + "</td>");
                            ExportHtml += "<td  colspan=2 style='text-align:right'>" + DisplayUtils.GetSystemAmountFormat(dLCRunningTotal) + "</td>";
                            Item["Bal_LocalCurrencyAmount"] = DisplayUtils.GetSystemAmountFormat(dLCRunningTotal);

                            if (dLCRunningTotal > 0)
                            {
                                strGLInquiryTable.Append("<td>Dr</td>");
                                ExportHtml += "<td>Dr</td>";
                            }
                            else if (dLCRunningTotal < 0)
                            {
                                strGLInquiryTable.Append("<td>Cr</td>");
                                ExportHtml += "<td>Dr</td>";
                            }
                            else
                            {
                                strGLInquiryTable.Append("<td>&nbsp;</td>");
                                ExportHtml += "<td>&nbsp;</td>";
                            }
                            #endregion
                        }
                        #endregion

                        strGLInquiryTable.Append("<td></td>");
                        strGLInquiryTable.Append("</tr>");

                        ExportHtml += "<td></td>";
                        ExportHtml += "</tr>";
                    }
                }
                else
                {
                    strGLInquiryTable.Append("<tr>");
                    strGLInquiryTable.Append("<td></td>");
                    strGLInquiryTable.Append("<td></td>");
                    strGLInquiryTable.Append("<td></td>");
                    strGLInquiryTable.Append("<td></td>");
                    strGLInquiryTable.Append("<td></td>");
                    strGLInquiryTable.Append("</tr>");
                }
                strGLInquiryTable.Append(GetGLTableFooter(pInquiriesViewModel, myAccountCode, dFCRunningTotal, dLCRunningTotal));
                pInquiriesViewModel.table = Common.toString(strGLInquiryTable);
                Session["tableHtml"] = ExportHtml;//Common.toString(strGLInquiryTable);
                Session["objGLInquiry"] = objGLInquiry;
                Session["InquiriesViewModel"] = pInquiriesViewModel;
            }
            else
            {
                pInquiriesViewModel.ErrMessage = Common.GetAlertMessage(1, pInquiriesViewModel.ErrMessage);
            }
            return pInquiriesViewModel;
        }

        protected HeaderViewModel GetGLTableHeader(string pAccountCode, InquiriesViewModel pInquiriesViewModel)
        {
            HeaderViewModel headerViewModel = new HeaderViewModel();
            string strGLTableHeader = "<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>";
            ExportHtml += "<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>";
            string LocalCurrencyCode = Common.GetSysSettings("DefaultCurrency");
            string ForeignCurrencyCode = "";
            if (pAccountCode.Trim() != "")
            {
                //string mAccountTitle = PostingUtils.getGLAccountNameByCode(pAccountCode.Trim());
                //GLChartOfAccounts objAccount = db.GLChartOfAccounts.Where(x => x.AccountCode == pAccountCode).SingleOrDefault();
                string sql = "select AccountName,CurrencyISOCode from GLChartOfAccounts where AccountCode='" + pAccountCode + "'";
                DataTable objAccount = DBUtils.GetDataTable(sql);
                if (objAccount != null && objAccount.Rows.Count > 0)
                {
                    strGLTableHeader += "<b>" + pAccountCode + " - " + Common.toString(objAccount.Rows[0]["AccountName"]) + "</b>";
                    ExportHtml += "<b>" + pAccountCode + " - " + Common.toString(objAccount.Rows[0]["AccountName"]) + " </b>";
                    ForeignCurrencyCode = Common.toString(objAccount.Rows[0]["CurrencyISOCode"]);
                }
                else
                {
                    strGLTableHeader += "<b>" + pAccountCode + " - </b>";
                    ExportHtml += "<b>" + pAccountCode + " - </b>";
                }
            }
            strGLTableHeader += "<table style='table-layout: auto; empty-cells: show;' border='1' rules='rows'>";
            strGLTableHeader += "<thead>";

            strGLTableHeader += "<tr>";
            strGLTableHeader += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 5%;'>V.No</th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 7%;'>Date</th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 25%;'>Description</th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 8%;'>Tranx #</th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 8%;'>Payout Amt</th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col'  colspan='2' style='text-align:center; border: 1px solid antiquewhite; width: 15%;'>Debit</th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col'  colspan='2' style='text-align:center; border: 1px solid antiquewhite; width: 15%;'>Credit</th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col'  colspan='2' style='text-align:center; border: 1px solid antiquewhite; width: 15%;'>Balance</th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col' style='width: 2%;'>&nbsp;</th>";
            strGLTableHeader += "</tr>";

            strGLTableHeader += "<tr>";
            strGLTableHeader += "<th class=' rgHeader' scope='col'></th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col'></th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col'></th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col'></th>";
            strGLTableHeader += "<th class=' rgHeader' scope='col'></th>";

            ExportHtml += "<table style='table-layout: auto; empty-cells: show; width: 100%;' border='1' rules='rows'>";
            ExportHtml += "<thead style='background-color: grey;'>";

            ExportHtml += "<tr>";
            ExportHtml += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 5%;'>V.No</th>";
            ExportHtml += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 7%;'>Date</th>";
            ExportHtml += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 25%;'>Description</th>";
            ExportHtml += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 8%;'>Tranx #</th>";
            ExportHtml += "<th class=' rgHeader' scope='col' style='border: 1px solid antiquewhite; width: 8%;'>Payout Amt</th>";
            ExportHtml += "<th class=' rgHeader' scope='col'  colspan='2' style='text-align:center; border: 1px solid antiquewhite; width: 15%;'>Debit</th>";
            ExportHtml += "<th class=' rgHeader' scope='col'  colspan='2' style='text-align:center; border: 1px solid antiquewhite; width: 15%;'>Credit</th>";
            ExportHtml += "<th class=' rgHeader' scope='col'  colspan='2' style='text-align:center; border: 1px solid antiquewhite; width: 15%;'>Balance</th>";
            ExportHtml += "<th class=' rgHeader' scope='col' style='width: 2%;'>&nbsp;</th>";
            ExportHtml += "</tr>";

            ExportHtml += "<tr>";
            ExportHtml += "<th class=' rgHeader' scope='col'></th>";
            ExportHtml += "<th class=' rgHeader' scope='col'></th>";
            ExportHtml += "<th class=' rgHeader' scope='col'></th>";
            ExportHtml += "<th class=' rgHeader' scope='col'></th>";
            ExportHtml += "<th class=' rgHeader' scope='col'></th>";
            if (pInquiriesViewModel.PdfCurrencyType == "0") // Both
            {
                strGLTableHeader += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + ForeignCurrencyCode + "</th>";
                strGLTableHeader += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + LocalCurrencyCode + "</th>";
                strGLTableHeader += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + ForeignCurrencyCode + "</th>";
                strGLTableHeader += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + LocalCurrencyCode + "</th>";
                strGLTableHeader += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + ForeignCurrencyCode + "</th>";
                strGLTableHeader += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + LocalCurrencyCode + "</th>";

                ExportHtml += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + ForeignCurrencyCode + "</th>";
                ExportHtml += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + LocalCurrencyCode + "</th>";
                ExportHtml += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + ForeignCurrencyCode + "</th>";
                ExportHtml += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + LocalCurrencyCode + "</th>";
                ExportHtml += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + ForeignCurrencyCode + "</th>";
                ExportHtml += "<th class=' rgHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>" + LocalCurrencyCode + "</th>";
            }
            else if (pInquiriesViewModel.PdfCurrencyType == "1") // Foreign
            {
                strGLTableHeader += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + ForeignCurrencyCode + "</th>";
                strGLTableHeader += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + ForeignCurrencyCode + "</th>";
                strGLTableHeader += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + ForeignCurrencyCode + "</th>";

                ExportHtml += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + ForeignCurrencyCode + "</th>";
                ExportHtml += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + ForeignCurrencyCode + "</th>";
                ExportHtml += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + ForeignCurrencyCode + "</th>";
            }
            else if (pInquiriesViewModel.PdfCurrencyType == "2") // Local
            {
                strGLTableHeader += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + LocalCurrencyCode + "</th>";
                strGLTableHeader += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + LocalCurrencyCode + "</th>";
                strGLTableHeader += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + LocalCurrencyCode + "</th>";

                ExportHtml += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + LocalCurrencyCode + "</th>";
                ExportHtml += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + LocalCurrencyCode + "</th>";
                ExportHtml += "<th class=' rgHeader' scope='col' colspan='2' style='text-align:right'>" + LocalCurrencyCode + "</th>";
            }
            strGLTableHeader += "<th class=' rgHeader' scope='col'>&nbsp;</th>";
            strGLTableHeader += "</tr>";
            strGLTableHeader += "</thead>";

            strGLTableHeader += "<tbody>";

            ExportHtml += "<th class=' rgHeader' scope='col'>&nbsp;</th>";
            ExportHtml += "</tr>";
            ExportHtml += "</thead>";

            ExportHtml += "<tbody>";

            if (pAccountCode.Trim() != "" && pInquiriesViewModel.PdfFromDate != null)
            {
                strGLTableHeader += "<tr style='background-color:rgb(255, 188, 0)'>";
                strGLTableHeader += "<td colspan='5' ><b>Opening Balance - " + pInquiriesViewModel.PdfFromDate.ToShortDateString() + ":</b></td>";

                ExportHtml += "<tr style='background-color:rgb(255, 188, 0)'>";
                ExportHtml += "<td colspan='5' ><b>Opening Balance - " + pInquiriesViewModel.PdfFromDate.ToShortDateString() + ":</b></td>";

                //string Sum_balances = PostingUtils.GetOpeningBalance(pAccountCode.Trim(), pInquiriesViewModel.FromDate);
                //string[] sum_balances_array = Sum_balances.Split(',');
                AccountBalance AcBalances = PostingUtils.GetOpeningBalance(pAccountCode.Trim(), pInquiriesViewModel.PdfFromDate);
                headerViewModel.FCOpeningBalance = AcBalances.ForeignCurrencyBalance;
                headerViewModel.LCOpeningBalance = AcBalances.LocalCurrencyBalance;

                if (pInquiriesViewModel.PdfCurrencyType == "0")
                {
                    #region Debit
                    if (headerViewModel.FCOpeningBalance > 0)
                    {
                        strGLTableHeader += "<td style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                        //Item["credit_LocalCurrencyAmount"] = Item["LocalCurrencyAmount"];
                        ExportHtml += "<td style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    if (headerViewModel.LCOpeningBalance > 0)
                    {
                        strGLTableHeader += "<td style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                        //Item["credit_LocalCurrencyAmount"] = Item["LocalCurrencyAmount"];
                        ExportHtml += "<td style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    #endregion

                    #region Credit
                    if (headerViewModel.FCOpeningBalance < 0)
                    {
                        strGLTableHeader += "<td style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    if (headerViewModel.LCOpeningBalance < 0)
                    {
                        strGLTableHeader += "<td style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    #endregion

                    #region Header
                    strGLTableHeader += "<td style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                    strGLTableHeader += "<td style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";

                    ExportHtml += "<td style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                    ExportHtml += "<td style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                    if (headerViewModel.LCOpeningBalance > 0)
                    {
                        strGLTableHeader += "<td>Dr</td>";
                        ExportHtml += "<td>Dr</td>";
                    }
                    else if (headerViewModel.LCOpeningBalance < 0)
                    {
                        strGLTableHeader += "<td>Cr</td>";
                        ExportHtml += "<td>Cr</td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    #endregion
                }
                else if (pInquiriesViewModel.PdfCurrencyType == "1") // Foreign
                {
                    if (headerViewModel.FCOpeningBalance > 0)
                    {
                        strGLTableHeader += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td colspan=2>&nbsp;</td>";
                        ExportHtml += "<td colspan=2>&nbsp;</td>";
                    }
                    if (headerViewModel.FCOpeningBalance < 0)
                    {
                        strGLTableHeader += "<td colspan=2>&nbsp;</td>";
                        ExportHtml += "<td colspan=2>&nbsp;</td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                    }
                    strGLTableHeader += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                    ExportHtml += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.FCOpeningBalance.ToString("0.00") + "</b></td>";
                    if (headerViewModel.FCOpeningBalance > 0)
                    {
                        strGLTableHeader += "<td>Dr</td>";
                        ExportHtml += "<td>Dr</td>";
                    }
                    else if (headerViewModel.FCOpeningBalance < 0)
                    {
                        strGLTableHeader += "<td>Cr</td>";
                        ExportHtml += "<td>Cr</td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                }
                else if (pInquiriesViewModel.PdfCurrencyType == "2") //Local
                {
                    if (headerViewModel.LCOpeningBalance > 0)
                    {
                        strGLTableHeader += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td colspan=2>&nbsp;</td>";
                        ExportHtml += "<td colspan=2>&nbsp;</td>";
                    }
                    if (headerViewModel.LCOpeningBalance < 0)
                    {
                        strGLTableHeader += "<td colspan=2>&nbsp;</td>";
                        ExportHtml += "<td colspan=2>&nbsp;</td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";

                    }
                    strGLTableHeader += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                    ExportHtml += "<td colspan=2 style='text-align:right'><b>" + headerViewModel.LCOpeningBalance.ToString("0.00") + "</b></td>";
                    if (headerViewModel.LCOpeningBalance > 0)
                    {
                        strGLTableHeader += "<td>Dr</td>";
                        ExportHtml += "<td>Dr</td>";
                    }
                    else if (headerViewModel.LCOpeningBalance < 0)
                    {
                        strGLTableHeader += "<td>Cr</td>";
                        ExportHtml += "<td>Cr</td>";
                    }
                    else
                    {
                        strGLTableHeader += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                }
                strGLTableHeader += "</tr>";
                ExportHtml += "</tr>";
            }

            headerViewModel.HtmlBody = strGLTableHeader;
            //return strGLTableHeader;
            return headerViewModel;
        }

        protected string GetGLTableFooter(InquiriesViewModel pInquiriesViewModel, string pAccountCode, decimal pFCRunningTotal, decimal pLCRunningTotal)
        {
            //InquiriesViewModel pInquiriesViewModel = new InquiriesViewModel();
            string strGLTableFooter = "";
            if (!string.IsNullOrEmpty(pAccountCode))
            {
                strGLTableFooter += "<tr style='background-color:rgb(255, 188, 0)'>";
                strGLTableFooter += "<td colspan='5' ><b>Ending Balance - " + pInquiriesViewModel.PdfToDate.ToShortDateString() + ":</b></td>";
                ExportHtml += "<tr style='background-color:rgb(255, 188, 0)'>";
                ExportHtml += "<td colspan='5' ><b>Ending Balance - " + pInquiriesViewModel.PdfToDate.ToShortDateString() + ":</b></td>";
                if (pInquiriesViewModel.PdfCurrencyType == "0")
                {
                    #region Debit
                    if (pFCRunningTotal > 0)
                    {
                        strGLTableFooter += "<td style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    if (pLCRunningTotal > 0)
                    {
                        strGLTableFooter += "<td style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    #endregion

                    #region Credit
                    if (pFCRunningTotal < 0)
                    {
                        strGLTableFooter += "<td style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    if (pLCRunningTotal < 0)
                    {
                        strGLTableFooter += "<td style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    #endregion

                    #region Header
                    strGLTableFooter += "<td style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                    strGLTableFooter += "<td style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                    ExportHtml += "<td style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                    ExportHtml += "<td style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                    if (pLCRunningTotal > 0)
                    {
                        strGLTableFooter += "<td>Dr</td>";
                        ExportHtml += "<td>Dr</td>";
                    }
                    else if (pLCRunningTotal < 0)
                    {
                        strGLTableFooter += "<td>Cr</td>";
                        ExportHtml += "<td>Cr</td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    #endregion
                }
                else if (pInquiriesViewModel.PdfCurrencyType == "1") //Foreign
                {
                    #region Debit
                    if (pFCRunningTotal > 0)
                    {
                        strGLTableFooter += "<td colspan=2 style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td colspan=2 style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td colspan=2>&nbsp;</td>";
                        ExportHtml += "<td colspan=2>&nbsp;</td>";
                    }
                    #endregion

                    #region Credit
                    if (pFCRunningTotal < 0)
                    {
                        strGLTableFooter += "<td colspan=2 style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td colspan=2 style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td colspan=2>&nbsp;</td>";
                        ExportHtml += "<td colspan=2>&nbsp;</td>";
                    }
                    #endregion

                    #region Header
                    strGLTableFooter += "<td colspan=2 style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                    ExportHtml += "<td colspan=2 style='text-align:right'><b>" + pFCRunningTotal.ToString("0.00") + "</b></td>";
                    if (pFCRunningTotal > 0)
                    {
                        strGLTableFooter += "<td>Dr</td>";
                        ExportHtml += "<td>Dr</td>";
                    }
                    else if (pFCRunningTotal < 0)
                    {
                        strGLTableFooter += "<td>Cr</td>";
                        ExportHtml += "<td>Cr</td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    #endregion
                }
                else if (pInquiriesViewModel.PdfCurrencyType == "2")
                {
                    #region Debit
                    if (pLCRunningTotal > 0)
                    {
                        strGLTableFooter += "<td colspan=2 style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td colspan=2 style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td colspan=2>&nbsp;</td>";
                        ExportHtml += "<td colspan=2>&nbsp;</td>";
                    }
                    #endregion

                    #region Credit
                    if (pLCRunningTotal < 0)
                    {
                        strGLTableFooter += "<td colspan=2 style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                        ExportHtml += "<td colspan=2 style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td colspan=2>&nbsp;</td>";
                        ExportHtml += "<td colspan=2>&nbsp;</td>";
                    }
                    #endregion

                    #region Header
                    strGLTableFooter += "<td colspan=2 style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                    ExportHtml += "<td colspan=2 style='text-align:right'><b>" + pLCRunningTotal.ToString("0.00") + "</b></td>";
                    if (pLCRunningTotal > 0)
                    {
                        strGLTableFooter += "<td>Dr</td>";
                        ExportHtml += "<td>Dr</td>";
                    }
                    else if (pLCRunningTotal < 0)
                    {
                        strGLTableFooter += "<td>Cr</td>";
                        ExportHtml += "<td>Cr</td>";
                    }
                    else
                    {
                        strGLTableFooter += "<td>&nbsp;</td>";
                        ExportHtml += "<td>&nbsp;</td>";
                    }
                    #endregion
                }

                strGLTableFooter += "</tr>";
                ExportHtml += "</tr>";
            }
            strGLTableFooter += "</tbody>";
            strGLTableFooter += "</table>";
            strGLTableFooter += "</div>";

            ExportHtml += "</tbody>";
            ExportHtml += "</table>";
            ExportHtml += "</div>";
            return strGLTableFooter;
        }

        public JsonResult GetGLAccounts([DataSourceRequest] DataSourceRequest request)
        {
            var bankAccount = PostingUtils.GetGeneralAccounts();
            return Json(bankAccount, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetDimensions([DataSourceRequest] DataSourceRequest request)
        {
            bool status = true;
            //var dimensions = db.GLDimensions.Where(x => x.Status == status).Select(c => new
            //{
            //    DimensionID = c.GLDimensionID,
            //    Title = c.Title
            //}).ToList();
            //dimensions.Insert(0, new { DimensionID = 0, Title = "All Branches" });

            GlDimensionModel obj = new GlDimensionModel();
            var dimensions = _dbcontext.Query<GlDimensionModel>("select GLDimensionID as DimensionID, Title as Title from GLDimensions where Status=1 order by Title").ToList();
            obj.Title = "All Branches";
            obj.DimensionID = 0;
            dimensions.Insert(0, obj);

            //LISTSUtils.PopulateDROPDOWN(ref ddlDimension1, "GLDimensionID", "Title", "", "select GLDimensionID, Title from GLDimensions where status=1", "Select Dimention");
            return Json(dimensions, JsonRequestBehavior.AllowGet);
        }
        //public void GLInquiryExportToPDF(InquiriesViewModel pInquiriesViewModel)
        //{
        //    string HostName = System.Web.HttpContext.Current.Request.Url.Host;
        //    if (pInquiriesViewModel.PdfCurrencyType == null)
        //    {
        //        pInquiriesViewModel.PdfCurrencyType = "0";
        //    }
        //    pInquiriesViewModel = BuildGLInquiryTable(pInquiriesViewModel);
        //    ////string fileName = Common.getRandomDigit(4) + "-" + "Ledger";
        //    string fileName = "Ledger-" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/");
        //    HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //    htmlToPdfConverter.BrowserWidth = 1200;
        //    // set HTML Load timeout
        //    htmlToPdfConverter.HtmlLoadedTimeout = 120;
        //    // set PDF page margins
        //    htmlToPdfConverter.Document.Margins = new PdfMargins(5, 5, 5, 20);
        //    htmlToPdfConverter.Document.PageSize = PdfPageSize.A4;
        //    htmlToPdfConverter.Document.PageOrientation = PdfPageOrientation.Portrait;
        //    htmlToPdfConverter.Document.PdfStandard = PdfStandard.Pdf;
        //    htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //    PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(pInquiriesViewModel.table, sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    Response.Clear();
        //    Response.ClearContent();
        //    Response.ClearHeaders();
        //    Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".pdf");
        //    Response.ContentType = "application/pdf";
        //    Response.Flush();
        //    Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".pdf"));
        //    Response.End();
        //}


        ////[HttpGet]
        ////public ActionResult DownloadExport(string ExportOption, string mode)
        ////{
        ////    string contentType = string.Empty;
        ////    string path = string.Empty;
        ////    string FileName = string.Empty;
        ////    try
        ////    {
        ////        if (mode == "action")
        ////        {
        ////            return Json(new { fileName = path }, JsonRequestBehavior.AllowGet);
        ////        }

        ////        path = Path.GetDirectoryName(Server.MapPath("~"));
        ////        path += "\\Export\\";
        ////        if (ExportOption == "PDF")
        ////        {
        ////            if (Session["tableHtml"] != null)
        ////            {
        ////                ExportHtml += Session["tableHtml"].ToString();
        ////                FileName = Common.getRandomString(6) + "-Ledger.pdf";
        ////                path += FileName;
        ////                HtmlToPdf htmlToPdfConverter = new HtmlToPdf();

        ////                htmlToPdfConverter.BrowserWidth = 1200;
        ////                // set HTML Load timeout
        ////                htmlToPdfConverter.HtmlLoadedTimeout = 120;
        ////                // set PDF page margins
        ////                htmlToPdfConverter.Document.Margins = new PdfMargins(5, 5, 5, 20);
        ////                htmlToPdfConverter.Document.PageSize = PdfPageSize.A4;
        ////                htmlToPdfConverter.Document.PageOrientation = PdfPageOrientation.Portrait;
        ////                htmlToPdfConverter.Document.PdfStandard = PdfStandard.Pdf;

        ////                htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        ////                //SetFooter(htmlToPdfConverter.Document);
        ////                byte[] pdfBuffer = htmlToPdfConverter.ConvertHtmlToMemory(ExportHtml, path);

        ////                System.IO.File.WriteAllBytes(path, pdfBuffer);

        ////            }
        ////        }
        ////        else if (ExportOption == "EXCEL")
        ////        {
        ////            FileName = Common.getRandomString(6) + "-Ledger.xlsx";
        ////            path += FileName;
        ////            ExportToExcel(path);
        ////        }
        ////        else
        ////        {
        ////            FileName = Common.getRandomString(6) + "-Ledger.csv";
        ////            path += FileName;
        ////            ExportToCSV(path);
        ////        }
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        return HttpNotFound();
        ////    }

        ////    if (!System.IO.File.Exists(path))
        ////    {
        ////        return HttpNotFound();
        ////    }

        ////    if (ExportOption == "PDF")
        ////    {
        ////        contentType = "application/pdf";
        ////    }
        ////    else if (ExportOption == "EXCEL")
        ////    {
        ////        contentType = "application/xlsx";
        ////    }
        ////    else
        ////    {
        ////        contentType = "text/plain";
        ////    }

        ////    return File(path, contentType, FileName);
        ////}
        //private void SetFooter(PdfDocumentControl htmlToPdfDocument)
        //{
        //    // enable footer display
        //    htmlToPdfDocument.Footer.Enabled = true;

        //    // set footer height
        //    htmlToPdfDocument.Footer.Height = 20;

        //    // set footer background color
        //    htmlToPdfDocument.Footer.BackgroundColor = System.Drawing.Color.WhiteSmoke;

        //    float pdfPageWidth =
        //            htmlToPdfDocument.PageOrientation == PdfPageOrientation.Portrait ?
        //            htmlToPdfDocument.PageSize.Width : htmlToPdfDocument.PageSize.Height;

        //    float footerWidth = pdfPageWidth - htmlToPdfDocument.Margins.Left - htmlToPdfDocument.Margins.Right;
        //    float footerHeight = htmlToPdfDocument.Footer.Height;

        //    //// layout HTML in footer
        //    //PdfHtml footerHtml = new PdfHtml(5, 5,
        //    //    @"<span style=""color:Navy; font-family:Times New Roman; font-style:italic"">
        //    //    Quickly Create High Quality PDFs with </span><a href=""http://www.hiqpdf.com"">HiQPdf</a>",
        //    //    null);
        //    //footerHtml.FitDestHeight = true;
        //    //htmlToPdfDocument.Footer.Layout(footerHtml);


        //    // add page numbering in a text element
        //    System.Drawing.Font pageNumberingFont =
        //        new System.Drawing.Font(new System.Drawing.FontFamily("Times New Roman"),
        //                    8, System.Drawing.GraphicsUnit.Point);
        //    PdfText pageNumberingText = new PdfText(5, footerHeight - 12,
        //            "Page {CrtPage} of {PageCount}", pageNumberingFont);
        //    pageNumberingText.HorizontalAlign = PdfTextHAlign.Center;
        //    pageNumberingText.EmbedSystemFont = true;
        //    pageNumberingText.ForeColor = System.Drawing.Color.DarkGreen;
        //    htmlToPdfDocument.Footer.Layout(pageNumberingText);


        //    //string footerImageFile = Server.MapPath("~") + @"\DemoFiles\Images\HiQPdfLogo.png";
        //    //PdfImage logoFooterImage = new PdfImage(footerWidth - 40 - 5, 5, 40,
        //    //            System.Drawing.Image.FromFile(footerImageFile));
        //    //htmlToPdfDocument.Footer.Layout(logoFooterImage);

        //    // create a border for footer
        //    //PdfRectangle borderRectangle = new PdfRectangle(1, 1, footerWidth - 2, footerHeight - 2);
        //    //borderRectangle.LineStyle.LineWidth = 0.5f;
        //    //borderRectangle.ForeColor = System.Drawing.Color.DarkGreen;
        //    //htmlToPdfDocument.Footer.Layout(borderRectangle);
        //}

        #region Export Helpers
        /// <summary>
        /// Method to export CSV File
        /// </summary>
        /// <param name="path"></param>
        public void ExportToCSV(string path)
        {
            InquiriesViewModel pInquiriesViewModel = new InquiriesViewModel();
            DataTable dt = null;
            try
            {
                if (Session["objGLInquiry"] != null)
                    dt = (DataTable)Session["objGLInquiry"];
                if (dt != null)
                {
                    StringBuilder sb = new StringBuilder();
                    IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                    sb.AppendLine(string.Join(",", columnNames));
                    foreach (DataRow row in dt.Rows)
                    {
                        IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                        sb.AppendLine(string.Join(",", fields));
                    }
                    File.WriteAllText(path, sb.ToString());
                }
                dt = null;
                dt.Dispose();
            }
            catch (Exception)
            {
                dt = null;
                dt.Dispose();
            }
        }

        /// <summary>
        /// Method to Export a file to Excel
        /// </summary>
        /// <param name="path"></param>
        public void ExportToExcel(InquiriesViewModel pInquiriesViewModel)
        {
            pInquiriesViewModel.PdfFromDate = pInquiriesViewModel.ExcelFromDate;
            pInquiriesViewModel.PdfToDate = pInquiriesViewModel.ExcelToDate;
            pInquiriesViewModel.PdfAccountCode = pInquiriesViewModel.ExcelAccountCode;
            pInquiriesViewModel.PdfCurrencyType = pInquiriesViewModel.ExcelCurrencyType;
            if (pInquiriesViewModel.ExcelCurrencyType == null)
            {
                pInquiriesViewModel.ExcelCurrencyType = "0";
                pInquiriesViewModel.PdfCurrencyType = "0";
            }

            pInquiriesViewModel = BuildGLInquiryTable(pInquiriesViewModel);

            DataTable objGLInquiry = null;
            try
            {
                //if (Session["InquiriesViewModel"] != null)
                //    pInquiriesViewModel = ((InquiriesViewModel)Session["InquiriesViewModel"]);
                //else
                //    return;
                if (Session["objGLInquiry"] != null)
                    objGLInquiry = (DataTable)Session["objGLInquiry"];
                else
                    return;
                int mycount1 = 11;

                string mySql1 = "select AccountName from GLChartOfAccounts where AccountCode='" + pInquiriesViewModel.ExcelAccountCode.ToString() + "'";
                //Response.Write("< mySql1 01 >> " + mySql1 + " <br/>");
                DataTable dt = DBUtils.GetDataTable(mySql1);
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                XLWorkbook workbook = new XLWorkbook();
                //string fileName = "Expo";
                #region Both 
                Response.Write("< ExcelCurrencyType >> " + pInquiriesViewModel.ExcelCurrencyType.ToString() + " <br/>");
                if (pInquiriesViewModel.ExcelCurrencyType.ToString() == "0")
                {

                    var worksheet = workbook.Worksheets.Add("Both Currency");
                    string currencyType = "Both";
                    string tbl_main = "<table Class='HTMLTableclass' style='width:100%;border-collapse:collapse;'>";
                    tbl_main += "<tr><td align='center' colspan='28' class='helpHed'>Both Currency</td></tr>";
                    worksheet.Cell("A1").Value = "Account Ledger - " + currencyType;
                    worksheet.Cell("A1").Style.Font.Bold = true;

                    worksheet.Cell("A5").Value = "Account Code : ";
                    worksheet.Cell("C5").Value = pInquiriesViewModel.ExcelAccountCode.ToString() + "          " + dt.Rows[0]["AccountName"].ToString();
                    worksheet.Cell("A7").Value = "Currency: ";
                    worksheet.Cell("C7").Value = objGLInquiry.Rows[0]["ForeignCurrencyISOCode"].ToString();
                    worksheet.Cell("I3").Value = "Print Date: ";
                    worksheet.Cell("I5").Value = "From Date: ";
                    worksheet.Cell("I7").Value = "To Date: ";
                    worksheet.Cell("K3").Value = DateTime.Now.ToShortDateString();
                    worksheet.Cell("K5").Value = pInquiriesViewModel.ExcelFromDate.ToShortDateString();
                    worksheet.Cell("K7").Value = pInquiriesViewModel.ExcelToDate.ToShortDateString();
                    worksheet.Cell("D9").Value = "Tranx #";
                    worksheet.Cell("E9").Value = "Payout Amount (in Case of Buyer)";
                    worksheet.Cell("F9").Value = "Debit";
                    worksheet.Cell("H9").Value = "Credit";
                    worksheet.Cell("J9").Value = "Balance";
                    worksheet.Cell("A10").Value = "V. No";
                    worksheet.Cell("B10").Value = "Date";
                    worksheet.Cell("C10").Value = "Description";
                    worksheet.Cell("F10").Value = objGLInquiry.Rows[0]["ForeignCurrencyISOCode"].ToString();
                    worksheet.Cell("G10").Value = SysPrefs.DefaultCurrency;
                    worksheet.Cell("H10").Value = objGLInquiry.Rows[0]["ForeignCurrencyISOCode"].ToString();
                    worksheet.Cell("I10").Value = SysPrefs.DefaultCurrency;
                    worksheet.Cell("J10").Value = objGLInquiry.Rows[0]["ForeignCurrencyISOCode"].ToString();
                    worksheet.Cell("K10").Value = SysPrefs.DefaultCurrency;
                    worksheet.Range("F9:G9").Merge();
                    worksheet.Range("H9:I9").Merge();
                    worksheet.Range("J9:K9").Merge();

                    worksheet.Cell("A9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("B9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("C9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("D9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("G9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("E9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("F9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("H9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("J9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("I9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("K9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("A10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("B10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("C10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("D10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("E10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("F10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("G10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("H10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("I10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("J10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("K10").Style.Fill.BackgroundColor = XLColor.LightGray;

                    worksheet.Cell("A9").Style.Font.Bold = true;
                    worksheet.Cell("B9").Style.Font.Bold = true;
                    worksheet.Cell("C9").Style.Font.Bold = true;
                    worksheet.Cell("D9").Style.Font.Bold = true;
                    worksheet.Cell("G9").Style.Font.Bold = true;
                    worksheet.Cell("E9").Style.Font.Bold = true;
                    worksheet.Cell("F9").Style.Font.Bold = true;
                    worksheet.Cell("H9").Style.Font.Bold = true;
                    worksheet.Cell("J9").Style.Font.Bold = true;
                    worksheet.Cell("I9").Style.Font.Bold = true;
                    worksheet.Cell("K9").Style.Font.Bold = true;
                    worksheet.Cell("A10").Style.Font.Bold = true;
                    worksheet.Cell("B10").Style.Font.Bold = true;
                    worksheet.Cell("C10").Style.Font.Bold = true;
                    worksheet.Cell("D10").Style.Font.Bold = true;
                    worksheet.Cell("E10").Style.Font.Bold = true;
                    worksheet.Cell("F10").Style.Font.Bold = true;
                    worksheet.Cell("G10").Style.Font.Bold = true;
                    worksheet.Cell("H10").Style.Font.Bold = true;
                    worksheet.Cell("I10").Style.Font.Bold = true;
                    worksheet.Cell("J10").Style.Font.Bold = true;
                    worksheet.Cell("K10").Style.Font.Bold = true;
                    worksheet.ColumnWidth = 15;
                    if (objGLInquiry != null)
                    {
                        foreach (DataRow my_row in objGLInquiry.Rows)
                        {
                            // Response.Write("< Row1 >>"+ objGLInquiry.Rows.Count+ "  <br/>");
                            worksheet.Cell("A" + mycount1).Value = my_row["VoucherNumber"].ToString();
                            worksheet.Cell("B" + mycount1).Value = Common.toDateTime(Common.toString(my_row["AddedDate"])).ToString("dd/MM/yyyy");
                            worksheet.Cell("C" + mycount1).Value = my_row["Memo"].ToString();
                            worksheet.Cell("D" + mycount1).Value = my_row["TransNo"].ToString();
                            worksheet.Cell("F" + mycount1).Value = my_row["debit_ForeignCurrencyAmount"].ToString();
                            worksheet.Cell("G" + mycount1).Value = my_row["debit_LocalCurrencyAmount"].ToString();
                            worksheet.Cell("H" + mycount1).Value = my_row["credit_ForeignCurrencyAmount"].ToString();
                            worksheet.Cell("I" + mycount1).Value = my_row["credit_LocalCurrencyAmount"].ToString();
                            worksheet.Cell("J" + mycount1).Value = my_row["Bal_ForeignCurrencyAmount"].ToString();
                            worksheet.Cell("K" + mycount1).Value = my_row["Bal_LocalCurrencyAmount"].ToString();
                            worksheet.Cell("L" + mycount1).Value = my_row["prefix"].ToString();
                            mycount1++;
                        }
                        worksheet.Cell("H" + mycount1).Value = objGLInquiry.Rows[objGLInquiry.Rows.Count - 1]["Bal_ForeignCurrencyAmount"].ToString();
                        worksheet.Cell("I" + mycount1).Value = objGLInquiry.Rows[objGLInquiry.Rows.Count - 1]["Bal_LocalCurrencyAmount"].ToString();
                        worksheet.Cell("J" + mycount1).Value = objGLInquiry.Rows[objGLInquiry.Rows.Count - 1]["Bal_ForeignCurrencyAmount"].ToString();
                        worksheet.Cell("K" + mycount1).Value = objGLInquiry.Rows[objGLInquiry.Rows.Count - 1]["Bal_LocalCurrencyAmount"].ToString();

                        worksheet.Cell("C" + mycount1).Value = "Total";
                        worksheet.Cell("F" + mycount1).Value = "0.00";
                        worksheet.Cell("G" + mycount1).Value = "0.00"; ;

                        worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("D" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("G" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("H" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("I" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("J" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("K" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;

                        worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("B" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("D" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("E" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("G" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("H" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("I" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("J" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("K" + mycount1).Style.Font.Bold = true;
                    }
                }
                #endregion

                #region FC/LC 
                else
                {

                    string currencyType = string.Empty;
                    if (pInquiriesViewModel.ExcelCurrencyType.ToString() == "0")
                        currencyType = "Both";
                    else if (pInquiriesViewModel.ExcelCurrencyType.ToString() == "1")
                        currencyType = "FC";
                    else
                        currencyType = "LC";
                    var worksheet = workbook.Worksheets.Add(currencyType + "Ledger");

                    string tbl_main = "<table Class='HTMLTableclass' style='width:100%;border-collapse:collapse;'>";
                    tbl_main += "<tr><td align='center' colspan='28' class='helpHed'>Both Currency</td></tr>";
                    worksheet.Cell("A1").Value = "Account Ledger - " + currencyType;
                    worksheet.Cell("A1").Style.Font.Bold = true;

                    worksheet.Cell("A5").Value = "Account Code : ";
                    worksheet.Cell("C5").Value = pInquiriesViewModel.ExcelAccountCode.ToString() + "          " + dt.Rows[0]["AccountName"].ToString();
                    worksheet.Cell("A7").Value = "Currency: ";
                    if (currencyType == "FC")
                        worksheet.Cell("C7").Value = objGLInquiry.Rows[0]["ForeignCurrencyISOCode"].ToString();
                    else
                        worksheet.Cell("C7").Value = SysPrefs.DefaultCurrency;
                    worksheet.Cell("G3").Value = "Print Date: ";
                    worksheet.Cell("G5").Value = "From Date: ";
                    worksheet.Cell("G7").Value = "To Date: ";
                    worksheet.Cell("H3").Value = DateTime.Now.ToShortDateString();
                    worksheet.Cell("H5").Value = pInquiriesViewModel.ExcelFromDate.ToShortDateString();
                    worksheet.Cell("H7").Value = pInquiriesViewModel.ExcelToDate.ToShortDateString();
                    worksheet.Cell("D9").Value = "Tranx #";
                    worksheet.Cell("E9").Value = "Payout Amount (in Case of Buyer)";
                    worksheet.Cell("F9").Value = "Debit";
                    worksheet.Cell("G9").Value = "Credit";
                    worksheet.Cell("H9").Value = "Balance";
                    worksheet.Cell("A10").Value = "V. No";
                    worksheet.Cell("B10").Value = "Date";
                    worksheet.Cell("C10").Value = "Description";
                    if (currencyType == "FC")
                    {
                        worksheet.Cell("F10").Value = pInquiriesViewModel.ExcelAccountCode.ToString();
                        worksheet.Cell("G10").Value = pInquiriesViewModel.ExcelAccountCode.ToString();
                        worksheet.Cell("H10").Value = pInquiriesViewModel.ExcelAccountCode.ToString();
                    }
                    else
                    {
                        worksheet.Cell("F10").Value = SysPrefs.DefaultCurrency;
                        worksheet.Cell("G10").Value = SysPrefs.DefaultCurrency;
                        worksheet.Cell("H10").Value = SysPrefs.DefaultCurrency;
                    }


                    worksheet.Cell("A9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("B9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("C9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("D9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("G9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("E9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("F9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("H9").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("A10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("B10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("C10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("D10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("E10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("F10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("G10").Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell("H10").Style.Fill.BackgroundColor = XLColor.LightGray;

                    worksheet.Cell("A9").Style.Font.Bold = true;
                    worksheet.Cell("B9").Style.Font.Bold = true;
                    worksheet.Cell("C9").Style.Font.Bold = true;
                    worksheet.Cell("D9").Style.Font.Bold = true;
                    worksheet.Cell("G9").Style.Font.Bold = true;
                    worksheet.Cell("E9").Style.Font.Bold = true;
                    worksheet.Cell("F9").Style.Font.Bold = true;
                    worksheet.Cell("H9").Style.Font.Bold = true;
                    worksheet.Cell("J9").Style.Font.Bold = true;
                    worksheet.Cell("I9").Style.Font.Bold = true;
                    worksheet.Cell("K9").Style.Font.Bold = true;
                    worksheet.Cell("A10").Style.Font.Bold = true;
                    worksheet.Cell("B10").Style.Font.Bold = true;
                    worksheet.Cell("C10").Style.Font.Bold = true;
                    worksheet.Cell("D10").Style.Font.Bold = true;
                    worksheet.Cell("E10").Style.Font.Bold = true;
                    worksheet.Cell("F10").Style.Font.Bold = true;
                    worksheet.Cell("G10").Style.Font.Bold = true;
                    worksheet.Cell("H10").Style.Font.Bold = true;

                    if (objGLInquiry != null)
                    {
                        foreach (DataRow my_row in objGLInquiry.Rows)
                        {
                            //Response.Write("< Row2 >>" + objGLInquiry.Rows.Count + "  <br/>");
                            //worksheet.Cell("A" + mycount1).Value = mycount1 - 11;
                            worksheet.Cell("A" + mycount1).Value = my_row["VoucherNumber"].ToString();
                            worksheet.Cell("B" + mycount1).Value = Common.toDateTime(Common.toString(my_row["AddedDate"])).ToString("dd/MM/yyyy"); ;
                            worksheet.Cell("C" + mycount1).Value = my_row["Memo"].ToString();
                            worksheet.Cell("D" + mycount1).Value = my_row["TransNo"].ToString();
                            if (currencyType == "FC")
                            {
                                worksheet.Cell("F" + mycount1).Value = my_row["debit_ForeignCurrencyAmount"].ToString();
                                worksheet.Cell("G" + mycount1).Value = my_row["credit_ForeignCurrencyAmount"].ToString();
                                worksheet.Cell("H" + mycount1).Value = my_row["Bal_ForeignCurrencyAmount"].ToString();
                            }
                            else
                            {
                                worksheet.Cell("F" + mycount1).Value = my_row["credit_LocalCurrencyAmount"].ToString();
                                worksheet.Cell("G" + mycount1).Value = my_row["debit_LocalCurrencyAmount"].ToString();
                                worksheet.Cell("H" + mycount1).Value = my_row["Bal_LocalCurrencyAmount"].ToString();
                            }
                            worksheet.Cell("I" + mycount1).Value = my_row["prefix"].ToString();
                            mycount1++;
                        }
                        if (currencyType == "FC")
                        {
                            worksheet.Cell("G" + mycount1).Value = objGLInquiry.Rows[objGLInquiry.Rows.Count - 1]["Bal_ForeignCurrencyAmount"].ToString();
                            worksheet.Cell("H" + mycount1).Value = objGLInquiry.Rows[objGLInquiry.Rows.Count - 1]["Bal_ForeignCurrencyAmount"].ToString();
                        }
                        else
                        {
                            worksheet.Cell("G" + mycount1).Value = objGLInquiry.Rows[objGLInquiry.Rows.Count - 1]["Bal_LocalCurrencyAmount"].ToString();
                            worksheet.Cell("H" + mycount1).Value = objGLInquiry.Rows[objGLInquiry.Rows.Count - 1]["Bal_LocalCurrencyAmount"].ToString();
                        }
                        worksheet.Cell("C" + mycount1).Value = "Total";
                        worksheet.Cell("F" + mycount1).Value = "0.00";

                        worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("D" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("G" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Cell("H" + mycount1).Style.Fill.BackgroundColor = XLColor.LightGray;

                        worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("B" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("D" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("E" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("G" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("H" + mycount1).Style.Font.Bold = true;

                    }

                }
                string FileName = Common.getRandomDigit(4) + "-" + "Ledger";
                string fileName = "Ledger-" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
                workbook.SaveAs(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".xlsx"));
                Response.Clear();
                Response.ClearHeaders();
                Response.ClearContent();
                Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".xlsx");
                Response.ContentType = "text/plain";
                Response.Flush();
                //Response.TransmitFile(Server.MapPath("~/Uploads/" + "/1/DeanRpt.xlsx"));
                Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".xlsx"));
                Response.End();
                #endregion
                //workbook.SaveAs(fileName);

                objGLInquiry = null;


                //objGLInquiry.Dispose();
            }
            catch (Exception ex)
            {
                objGLInquiry = null;
                Response.Write("<font color='Red'>" + ex.Message + " </font>");
            }
        }


        ///// <summary>
        ///// Method To Export Files depending upong Type Pass
        ///// </summary>
        ///// <param name="ExportOption">Type of file</param>
        ///// <returns></returns>
        //[HttpPost]
        //public JsonResult Export(string ExportOption)
        //{
        //    string FilePath = string.Empty;
        //    string path = string.Empty;
        //    string FileName = string.Empty;
        //    try
        //    {
        //        path = Path.GetDirectoryName(Server.MapPath("~"));
        //        path += "\\Export\\";
        //        if (ExportOption == "PDF")
        //        {
        //            if (Session["tableHtml"] != null)
        //            {
        //                ExportHtml += Session["tableHtml"].ToString();

        //                path += Common.getRandomString(6) + "-Ledger.pdf";
        //                HtmlToPdf htmlToPdfConverter = new HtmlToPdf();

        //                htmlToPdfConverter.BrowserWidth = 0;
        //                // set HTML Load timeout
        //                htmlToPdfConverter.HtmlLoadedTimeout = 120;
        //                // set PDF page margins
        //                htmlToPdfConverter.Document.Margins = new PdfMargins(5);
        //                htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //                byte[] pdfBuffer = htmlToPdfConverter.ConvertHtmlToMemory(ExportHtml, path);

        //                System.IO.File.WriteAllBytes(path, pdfBuffer);
        //                DownloadFile(ExportOption, path);
        //            }
        //        }
        //        else if (ExportOption == "EXCEL")
        //        {
        //            path += Common.getRandomString(6) + "-Ledger.xlsx";
        //            ExportToExcel(path);
        //            DownloadFile(ExportOption, path);

        //        }
        //        else
        //        {
        //            path += Common.getRandomString(6) + "-Ledger.csv";
        //            ExportToCSV(path);
        //            DownloadFile(ExportOption, path);
        //        }

        //    }
        //    catch (Exception)
        //    { }
        //    return Json(FilePath, JsonRequestBehavior.AllowGet);
        //}
        ///// <summary>
        ///// Function to Download File
        ///// </summary>
        ///// <param name="ExportOption">Provide the filetype you want to download</param>
        //public void DownloadFile(string ExportOption, string path)
        //{
        //    try
        //    {
        //        // Doawload file from server

        //        Response.Clear();
        //        Response.ClearHeaders();
        //        Response.ClearContent();
        //        #region Excel
        //        if (ExportOption == "EXCEL")
        //        {
        //            Response.AddHeader("Content-Disposition", "attachment; filename=Ledger.xlsx");
        //            Response.ContentType = "text/plain";
        //            Response.Flush();
        //            //Response.TransmitFile(Server.MapPath(Request.Url.GetLeftPart(UriPartial.Authority) + Request.ApplicationPath + "/Export/ledger.xlsx"));
        //            Response.TransmitFile(path);
        //        }
        //        #endregion
        //        #region PDF
        //        else if (ExportOption == "PDF")
        //        {
        //            Response.AddHeader("Content-Disposition", "attachment; filename=Ledger.pdf");

        //            Response.ContentType = "text/plain";
        //            Response.Flush();
        //            //Response.TransmitFile(Server.MapPath(Request.Url.GetLeftPart(UriPartial.Authority) + Request.ApplicationPath + "/Export/ledger.pdf"));
        //            Response.TransmitFile(path);
        //        }
        //        #endregion
        //        #region Excel
        //        else
        //        {
        //            Response.AddHeader("Content-Disposition", "attachment; filename=Ledger.csv");
        //            Response.ContentType = "text/plain";
        //            Response.Flush();
        //            //Response.TransmitFile(Server.MapPath(Request.Url.GetLeftPart(UriPartial.Authority) + Request.ApplicationPath + "/Export/ledger.csv"));
        //            Response.TransmitFile(path);
        //        }
        //        #endregion
        //        Response.End();
        //    }
        //    catch (Exception ex)
        //    { }
        //}

        #endregion
        #endregion

        #region Search Payment
        public ActionResult SearchPayment()
        {
            SearchPaymentViewModel pModel = new SearchPaymentViewModel();
            return View("~/Views/Inquiries/SearchPayment.cshtml", pModel);
        }
        public ActionResult PaymentView(string id)
        {
            SearchPaymentViewModel pGLTransViewModel = new SearchPaymentViewModel();
            if (!string.IsNullOrEmpty(id))
            {
                string Records = "Select PaymentNo, HoldDate, CancelledDate, SendingCountry,RecevingCountry,Address,City,TransactionID,Phone,Date,Recipient,SenderName,FC_Amount,CancellationReason,(select top 1 CustomerPrefix from Customers where Customers.CustomerID=CustomerTransactions.CustomerID ) as  CustomerPrefix,Status,PaidDate,Pounds  from CustomerTransactions Where TransactionID=" + id + "";
                DataTable transTable = DBUtils.GetDataTable(Records);
                if (transTable != null && transTable.Rows.Count > 0)
                {
                    foreach (DataRow RecordTable in transTable.Rows)
                    {
                        string getVoucherNo = "SELECT distinct GLReferences.VoucherNumber, GLTransactions.GLReferenceID FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID where TransactionNumber = '" + Common.toString(RecordTable["PaymentNo"]) + "'";
                        DataTable dtVouchers = DBUtils.GetDataTable(getVoucherNo);
                        if (dtVouchers != null && dtVouchers.Rows.Count > 0)
                        {
                            foreach (DataRow drVoucher in dtVouchers.Rows)
                            {
                                //pGLTransViewModel.VoucherNo += Common.toString(drVoucher["VoucherNumber"]) + ", ";
                                pGLTransViewModel.VoucherNo += "<a href=\"javascript: popup('/Inquiries/GLTransView?Id=" + Security.EncryptQueryString(drVoucher["GLReferenceID"].ToString()) + "')\")>" + Common.toString(drVoucher["VoucherNumber"]) + "</a> " + ", ";
                            }
                        }
                        string getAccountName = "select AccountName, AccountCode from GLChartOfAccounts where Prefix = '" + Common.toString(RecordTable["CustomerPrefix"]) + "'";
                        DataTable getHoldAcc = DBUtils.GetDataTable(getAccountName);
                        if (getHoldAcc != null && getHoldAcc.Rows.Count > 0)
                        {
                            foreach (DataRow drRow in getHoldAcc.Rows)
                            {
                                pGLTransViewModel.AccountName = Common.toString(drRow["AccountName"]);
                                pGLTransViewModel.HoldAccount = Common.toString(drRow["AccountCode"]);
                            }
                        }
                        pGLTransViewModel.CustomerPrefix = Common.toString(RecordTable["CustomerPrefix"]);
                        pGLTransViewModel.Phone = Common.toString(RecordTable["Phone"]);
                        pGLTransViewModel.PaymentNo = Common.toString(RecordTable["PaymentNo"]);
                        pGLTransViewModel.InvoiceNo = Common.toString(RecordTable["TransactionID"]);
                        pGLTransViewModel.Date = Common.toDateTime(RecordTable["Date"]);
                        pGLTransViewModel.CancelledDate = Common.toDateTime(RecordTable["CancelledDate"]);
                        pGLTransViewModel.Recipient = Common.toString(RecordTable["Recipient"]) + "<br> " + Common.toString(RecordTable["Address"] + "-" + Common.toString(RecordTable["City"]));
                        pGLTransViewModel.SenderName = Common.toString(RecordTable["SenderName"] + "<br> " + RecordTable["SendingCountry"]);
                        pGLTransViewModel.FC_Amount = Common.toDecimal(RecordTable["FC_Amount"]);
                        pGLTransViewModel.CancellationReason = Common.toString(RecordTable["CancellationReason"]);
                        pGLTransViewModel.Status = Common.toString(RecordTable["Status"]);
                        pGLTransViewModel.PaidDate = Common.toDateTime(RecordTable["PaidDate"]);
                        pGLTransViewModel.HoldDate = Common.toDateTime(RecordTable["HoldDate"]);
                        pGLTransViewModel.Pounds = Common.toDecimal(RecordTable["Pounds"]);
                    }
                }
            }
            return View("~/Views/Inquiries/PaymentView.cshtml", pGLTransViewModel);
        }
        public ActionResult PaymentSearchRead([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM CustomerTransactions /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                         CustomerTransactionID, PostingDate, PaymentNo, AgentPrefix, SenderName,Recipient,Phone,Status,Currency,FC_Amount,TransactionID,CustomerID,RecevingCountry,Pounds FROM CustomerTransactions
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";

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
                builder.Where("CustomerTransactionID = 0");
            }

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("CustomerTransactionID desc");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);
            var rows = _dbcontext.Query<SearchPaymentViewModel>(selector.RawSql, selector.Parameters);
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }

        #endregion

        #region Search Voucher
        public ActionResult SearchVoucher()
        {
            JournalInquiryViewModel pJournalInquiryViewModel = new JournalInquiryViewModel();
            return View("~/Views/Inquiries/SearchVoucher.cshtml", pJournalInquiryViewModel);
        }

        [HttpPost]
        public ActionResult SearchVoucher(JournalInquiryViewModel pJournalInquiryViewModel)
        {
            pJournalInquiryViewModel = SearchVoucherTable(pJournalInquiryViewModel);
            return View("~/Views/Inquiries/SearchVoucher.cshtml", pJournalInquiryViewModel);
        }
        public static JournalInquiryViewModel SearchVoucherTable(JournalInquiryViewModel pJournalInquiryViewModel)
        {
            pJournalInquiryViewModel.haveData = false;
            decimal cpTotalDebit = 0m;
            decimal cpTotalCredit = 0m;
            if (!string.IsNullOrEmpty(pJournalInquiryViewModel.VoucherNumber))
            {
                StringBuilder strMyTable = new StringBuilder();
                strMyTable.Append("<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>");
                strMyTable.Append("<table style='table-layout: auto; empty-cells: show;' border='1' rules='rows'>");
                strMyTable.Append("<thead>");
                strMyTable.Append("<tr>");
                strMyTable.Append("<th class='tabelHeader' align='center'>Date</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>V.No</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>Description</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>Account Title</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>Curr</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>F/C</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>Debit</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>Credit</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>P.By</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>A.By</th>");
                strMyTable.Append("<th class='tabelHeader' align='center'>Type</th>");
                strMyTable.Append("</tr>");
                strMyTable.Append("</thead>");
                strMyTable.Append("<tbody>");
                string mySql = "Select GLTransactions.GLReferenceID,  (Select AccountName From GLChartOfAccounts Where GLTransactions.AccountCode = GLChartOfAccounts.AccountCode) As AccountTitle, GLTransactions.AddedDate, GLTransactions.ForeignCurrencyAmount, GLTransactions.LocalCurrencyAmount, GLReferences.Memo, GLReferences.VoucherNumber, GLTransactions.TransactionNumber As TransNo, GLTransactions.ForeignCurrencyISOCode, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.PostedBy) as PostedBy, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.AuthorizedBy) as AuthBy,(select Description from ConfigItemsData where GLReferences.TypeiCode= ConfigItemsData.ItemsDataCode) as TypeiCode FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID Where TransactionID > 0";
                if (!string.IsNullOrEmpty(pJournalInquiryViewModel.VoucherNumber))
                {
                    mySql += " and GLReferences.VoucherNumber='" + pJournalInquiryViewModel.VoucherNumber + "'";
                }
                if (!pJournalInquiryViewModel.ShowAll)
                {
                    mySql += " and GLReferences.isPosted = 1";
                }
                mySql += " order by AddedDate Desc";
                DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
                if (dtGLTransactions != null)
                {
                    foreach (DataRow MyRow in dtGLTransactions.Rows)
                    {
                        pJournalInquiryViewModel.haveData = true;
                        strMyTable.Append("<tr>");
                        strMyTable.Append("<td align='left'>" + Common.toDateTime(MyRow["AddedDate"]).ToString("yyyy-MM-dd") + "</td>");

                        strMyTable.Append("<td align='left'>" + "<a href=\"javascript: showDetails('/Inquiries/GLTransView?Id=" + Security.EncryptQueryString(MyRow["GLReferenceID"].ToString()) + "','Voucher View')\")>" + MyRow["VoucherNumber"].ToString() + "</a></td>");
                        strMyTable.Append("<td align='left'>" + MyRow["Memo"].ToString() + "</td>");
                        strMyTable.Append("<td align='left'>" + MyRow["AccountTitle"].ToString() + "</td>");
                        strMyTable.Append("<td align='left'>" + MyRow["ForeignCurrencyISOCode"].ToString() + "</td>");
                        strMyTable.Append("<td align='center'>" + DisplayUtils.GetSystemAmountFormat(Math.Abs(Common.toDecimal(MyRow["ForeignCurrencyAmount"]))) + "</td>");
                        decimal LCAmount = Common.toDecimal(MyRow["LocalCurrencyAmount"]);
                        LCAmount = DisplayUtils.getRoundDigit(Common.toString(MyRow["LocalCurrencyAmount"]));
                        //if (Common.toString(MyRow["AccountTitle"]).Trim() == "Revaluation Gain & Loss")
                        //{
                        //    if (LCAmount >= 0)
                        //    {
                        //        LCAmount = -1*LCAmount;
                        //    }

                        //}

                        if (LCAmount >= 0)
                        {
                            strMyTable.Append("<td align='left'>" + DisplayUtils.GetSystemAmountFormat(LCAmount) + "</td>");
                            strMyTable.Append("<td align='left'>&nbsp;</td>");
                            cpTotalDebit = cpTotalDebit + LCAmount;
                        }
                        else
                        {
                            strMyTable.Append("<td align='left'>&nbsp;</td>");
                            strMyTable.Append("<td align='left'>" + DisplayUtils.GetSystemAmountFormat(Math.Abs(LCAmount)) + "</td>");
                            cpTotalCredit = cpTotalCredit + -1 * LCAmount;
                        }

                        strMyTable.Append("<td align='left'>" + MyRow["PostedBy"].ToString() + "</td>");
                        strMyTable.Append("<td align='left'>" + MyRow["AuthBy"].ToString() + "</td>");
                        strMyTable.Append("<td align='left'>" + MyRow["TypeiCode"].ToString() + "</td>");
                        strMyTable.Append("</tr>");
                    }
                }
                strMyTable.Append("<tr>");
                strMyTable.Append("<td colspan='4'></td>");
                strMyTable.Append("<td style='background-color:#1c3a70; color:#FFF;'>Total:</td>");
                strMyTable.Append("<td align='left'>&nbsp;</td>");
                strMyTable.Append("<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + DisplayUtils.GetSystemAmountFormat(cpTotalDebit) + "</td>");
                strMyTable.Append("<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + DisplayUtils.GetSystemAmountFormat(cpTotalCredit) + "</td>");
                strMyTable.Append("<td align='left'>&nbsp;</td>");
                strMyTable.Append("<td align='left'>&nbsp;</td>");
                strMyTable.Append("<td align='left'>&nbsp;</td>");
                strMyTable.Append("</tr>");
                strMyTable.Append("</tbody></table></div>");
                pJournalInquiryViewModel.Table = strMyTable.ToString();
            }
            else
            {
                pJournalInquiryViewModel.ErrMessage = Common.GetAlertMessage(1, "Please enter voucher no.");
            }
            return pJournalInquiryViewModel;

            //ltrTableBody.Text = Convert.ToString(strMyTable);
        }
        //public void VoucherExportToPDF(JournalInquiryViewModel pJournalInquiryViewModel)
        //{

        //    string HostName = System.Web.HttpContext.Current.Request.Url.Host;
        //    pJournalInquiryViewModel = SearchVoucherTable(pJournalInquiryViewModel);
        //    ////string fileName = Common.getRandomDigit(4) + "-" + "Voucher";
        //    string fileName = "Voucher-" + pJournalInquiryViewModel.VoucherNumber + "-" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/");
        //    HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //    htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //    PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(pJournalInquiryViewModel.Table, sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    Response.Clear();
        //    Response.ClearContent();
        //    Response.ClearHeaders();
        //    Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".pdf");
        //    Response.ContentType = "application/pdf";
        //    Response.Flush();
        //    Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".pdf"));
        //    Response.End();
        //}

        public void VoucherExportToExcel(JournalInquiryViewModel pJournalInquiryViewModel)
        {
            try
            {
                decimal cpTotalDebit = 0m;
                decimal cpTotalCredit = 0m;
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                XLWorkbook workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Daily Transactions List");
                worksheet.Cell("A1").Value = @SysPrefs.SiteName;
                worksheet.Cell("A1").Style.Font.Bold = true;
                worksheet.Cell("A2").Value = "Daily Transactions List";
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("E2").Value = "As on Dated : " + @SysPrefs.PostingDate + "";
                worksheet.Cell("E2").Style.Font.Bold = true;
                worksheet.Cell("I1").Value = "Print Date Time : " + DateTime.Now + "";
                worksheet.Cell("I1").Style.Font.Bold = true;

                worksheet.Cell("A3").Value = "Date";
                worksheet.Cell("B3").Value = "V.No";
                worksheet.Cell("C3").Value = "Description";
                worksheet.Cell("D3").Value = "Account Title";
                worksheet.Cell("E3").Value = "Curr";
                worksheet.Cell("F3").Value = "F/C Balance";
                worksheet.Cell("G3").Value = "Debit";
                worksheet.Cell("H3").Value = "Credit";
                worksheet.Cell("I3").Value = "P.By";
                worksheet.Cell("J3").Value = "A.By";
                worksheet.Cell("K3").Value = "Type";

                worksheet.Cell("A3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("B3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("C3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("D3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("E3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("F3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("G3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("H3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("I3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("J3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("K3").Style.Fill.BackgroundColor = XLColor.PersianRed;

                worksheet.Cell("A3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("B3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("C3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("D3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("E3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("F3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("G3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("H3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("I3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("J3").Style.Font.FontColor = XLColor.Black;
                worksheet.Cell("K3").Style.Font.FontColor = XLColor.Black;
                worksheet.ColumnWidth = 15;
                int mycount1 = 4;
                string mySql = "Select GLTransactions.GLReferenceID,  (Select AccountName From GLChartOfAccounts Where GLTransactions.AccountCode = GLChartOfAccounts.AccountCode) As AccountTitle, GLTransactions.AddedDate, GLTransactions.ForeignCurrencyAmount, GLTransactions.LocalCurrencyAmount, GLReferences.Memo, GLReferences.VoucherNumber, GLTransactions.TransactionNumber As TransNo, GLTransactions.ForeignCurrencyISOCode, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.PostedBy) as PostedBy, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.AuthorizedBy) as AuthBy,(select Description from ConfigItemsData where GLReferences.TypeiCode= ConfigItemsData.ItemsDataCode) as TypeiCode FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID Where TransactionID > 0";
                if (!string.IsNullOrEmpty(pJournalInquiryViewModel.VoucherNumber))
                {
                    mySql += " and GLReferences.VoucherNumber='" + pJournalInquiryViewModel.VoucherNumber + "'";
                }
                if (!pJournalInquiryViewModel.ShowAll)
                {
                    mySql += " and GLReferences.isPosted = 1";
                }
                mySql += " order by AddedDate Desc";
                DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
                if (dtGLTransactions != null)
                {
                    foreach (DataRow MyRow in dtGLTransactions.Rows)
                    {
                        worksheet.Cell("A" + mycount1).Value = Common.toDateTime(MyRow["AddedDate"]).ToString("yyyy-MM-dd");
                        worksheet.Cell("B" + mycount1).Value = MyRow["GLReferenceID"].ToString();
                        worksheet.Cell("C" + mycount1).Value = MyRow["Memo"].ToString();
                        worksheet.Cell("D" + mycount1).Value = MyRow["AccountTitle"].ToString();
                        worksheet.Cell("E" + mycount1).Value = MyRow["ForeignCurrencyISOCode"].ToString();
                        worksheet.Cell("F" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(Math.Abs(Common.toDecimal(MyRow["ForeignCurrencyAmount"])));
                        decimal LCAmount = Common.toDecimal(MyRow["LocalCurrencyAmount"]);
                        LCAmount = DisplayUtils.getRoundDigit(Common.toString(MyRow["LocalCurrencyAmount"]));
                        if (Common.toString(MyRow["AccountTitle"]).Trim() == "Revaluation Gain & Loss")
                        {
                            if (LCAmount >= 0)
                            {
                                LCAmount = -1 * LCAmount;
                            }
                        }
                        //if (Common.toString(MyRow["AccountTitle"]).Trim() == "Inc-Revaluation")
                        //{
                        //    if (LCAmount >= 0)
                        //    {
                        //        LCAmount = -1 * LCAmount;
                        //    }
                        //}
                        if (LCAmount >= 0)
                        {
                            worksheet.Cell("G" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(LCAmount);
                            worksheet.Cell("H" + mycount1).Value = "";
                            cpTotalDebit = cpTotalDebit + LCAmount;
                        }
                        else
                        {
                            worksheet.Cell("G" + mycount1).Value = "";
                            worksheet.Cell("H" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(Math.Abs(LCAmount));
                            cpTotalCredit = cpTotalCredit + -1 * LCAmount;
                        }
                        worksheet.Cell("I" + mycount1).Value = MyRow["PostedBy"].ToString();
                        worksheet.Cell("J" + mycount1).Value = MyRow["AuthBy"].ToString();
                        worksheet.Cell("K" + mycount1).Value = MyRow["TypeiCode"].ToString();
                        worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("D" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("G" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("H" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("I" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("J" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("K" + mycount1).Style.Fill.BackgroundColor = XLColor.White;

                        worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("B" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("D" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("E" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("G" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("H" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("I" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("J" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("K" + mycount1).Style.Font.Bold = true;
                        mycount1++;
                    }
                    worksheet.Cell("E" + mycount1).Value = "Total";
                    worksheet.Cell("G" + mycount1).Value = cpTotalDebit;
                    worksheet.Cell("H" + mycount1).Value = "(" + cpTotalCredit + ")";

                    worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("G" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("H" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;

                    worksheet.Cell("E" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("G" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("H" + mycount1).Style.Font.Bold = true;

                }

                //// string fileName = Common.getRandomDigit(4) + "-" + "Voucher";
                string fileName = "Voucher-" + pJournalInquiryViewModel.VoucherNumber + "-" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
                if (!Directory.Exists("/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp")) Directory.CreateDirectory("/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp");
                workbook.SaveAs(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".xlsx"));
                Response.Clear();
                Response.ClearHeaders();
                Response.ClearContent();
                Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".xlsx");
                Response.ContentType = "text/plain";
                Response.Flush();
                //Response.TransmitFile(Server.MapPath("~/Uploads/" + "/1/DeanRpt.xlsx"));
                Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".xlsx"));
                Response.End();
            }
            catch (Exception x)
            {
                Response.Write("<font color='Red'>" + x.Message + " </font>");

            }
        }
        #endregion

    }
    public class HeaderViewModel
    {
        public string HtmlBody { get; set; }
        public decimal FCOpeningBalance { get; set; }
        public decimal LCOpeningBalance { get; set; }

    }
}