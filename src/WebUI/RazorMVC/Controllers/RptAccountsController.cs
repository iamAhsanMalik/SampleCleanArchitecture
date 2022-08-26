using DHAAccounts.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using ClosedXML.Excel;
using System.Data.SqlClient;
using Kendo.Mvc.UI;
using Dapper;
using Kendo.Mvc;
using System.Globalization;

namespace Accounting.Controllers
{
    public class RptAccountsController : Controller
    {
        private AppDbContext _dbcontext = new AppDbContext();
        ////private AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());

        #region Journal Inquiry
        public ActionResult JournalInquiry()
        {
            JournalInquiryViewModel pJournalInquiryViewModel = new JournalInquiryViewModel();
            pJournalInquiryViewModel.FromDate = Common.toDateTime(SysPrefs.PostingDate);
            pJournalInquiryViewModel.ToDate = Common.toDateTime(SysPrefs.PostingDate);
            return View("~/Views/Reports/Accounts/journalInquiry.cshtml", pJournalInquiryViewModel);
        }

        public ActionResult CheckList_Read([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID  /**where**/";
            string selectQuery = @"SELECT  * FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum";
            selectQuery += @",GLTransactions.GLReferenceID,  (Select AccountName From GLChartOfAccounts Where GLTransactions.AccountCode = GLChartOfAccounts.AccountCode)";
            selectQuery += @"As AccountTitle, GLTransactions.AddedDate, GLTransactions.ForeignCurrencyAmount, GLTransactions.LocalCurrencyAmount, GLTransactions.Memo as Memo, ";
            selectQuery += @"GLReferences.VoucherNumber , GLTransactions.TransactionNumber As TransNo, GLTransactions.ForeignCurrencyISOCode, (select FirstName + ' ' + LastName from [Users]";
            selectQuery += @"where [Users].UserID= GLReferences.PostedBy) as PostedBy, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.AuthorizedBy) as AuthBy";
            selectQuery += @",(select Description from ConfigItemsData where GLReferences.TypeiCode= ConfigItemsData.ItemsDataCode) as TypeiCode FROM GLTransactions INNER JOIN GLReferences";
            selectQuery += @" ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID";
            selectQuery += @" /**where**/ ) AS RowConstrainedResult WHERE   RowNum >= (@PageIndex * @PageSize + 1 ) AND RowNum <= (@PageIndex + 1) * @PageSize ORDER BY RowNum";

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
            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("VoucherNumber Desc");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);
            var rows = _dbcontext.Query<JournalInquiryViewModel>(selector.RawSql, selector.Parameters);
            string voucherNo = "";
            foreach (JournalInquiryViewModel model in rows)
            {
                //if (voucherNo != model.VoucherNumber)
                //{
                // voucherNo = model.VoucherNumber;
                model.GLReferenceID = Security.EncryptQueryString(model.GLReferenceID).ToString();
                model.VoucherNumber = model.VoucherNumber;
                model.AddedDate = Common.toDateTime(model.AddedDate).ToShortDateString();
                model.Memo = model.Memo;
                //}
                //else
                //{
                //    model.VoucherNumber = "";
                //    model.AddedDate = "";
                //    model.Memo = "";
                //}
            }
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }

        #region Old Code

        [HttpPost]
        public ActionResult JournalInquiry(JournalInquiryViewModel pJournalInquiryViewModel)
        {
            pJournalInquiryViewModel = JournalInquiryTable(pJournalInquiryViewModel);
            return View("~/Views/Reports/Accounts/journalInquiry.cshtml", pJournalInquiryViewModel);
        }
        public static JournalInquiryViewModel JournalInquiryTable(JournalInquiryViewModel pJournalInquiryViewModel)
        {
            pJournalInquiryViewModel.haveData = false;
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

            //string mySql = "Select TransactionID, Convert(nvarchar(12), AddedDate, 101) as AddedDate, (Select Description From ConfigItemsData Where ConfigItemsData.ItemsDataCode = GLTransactions.AccountTypeiCode) As AccountTypeiCode, GLReferenceID, (Select ReferenceNo from GLReferences Where GLReferences.GLReferenceID= GLTransactions.GLReferenceID	) as ReferenceNo, isNull(LocalAmount, 0) as LocalAmount, Memo, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLTransactions.AddedBy) as AddedBy from GLTransactions Where TransactionID > 0";
            string mySql = "Select GLTransactions.GLReferenceID,  (Select AccountName From GLChartOfAccounts Where GLTransactions.AccountCode = GLChartOfAccounts.AccountCode) As AccountTitle, GLTransactions.AddedDate, GLTransactions.ForeignCurrencyAmount, GLTransactions.LocalCurrencyAmount, GLTransactions.Memo, GLReferences.VoucherNumber, GLTransactions.TransactionNumber As TransNo, GLTransactions.ForeignCurrencyISOCode, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.PostedBy) as PostedBy, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.AuthorizedBy) as AuthBy,(select Description from ConfigItemsData where GLReferences.TypeiCode= ConfigItemsData.ItemsDataCode) as TypeiCode FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID Where TransactionID > 0";
            //GLReferences.isPosted = 1 and GLTransactions.AccountCode = '" + myAccountCode + "'";
            if (!string.IsNullOrEmpty(pJournalInquiryViewModel.VoucherNumber))
            {
                mySql += " and GLReferences.VoucherNumber='" + pJournalInquiryViewModel.VoucherNumber + "'";
            }
            if (!string.IsNullOrEmpty(pJournalInquiryViewModel.Memo))
            {
                mySql += " and GLReferences.Memo  like '%" + pJournalInquiryViewModel.Memo + "%'";
            }
            if (pJournalInquiryViewModel.PdfFromDate != null)
            {
                mySql += " and AddedDate >='" + pJournalInquiryViewModel.PdfFromDate.ToString("yyyy-MM-dd") + " 00:00:00' ";
            }
            if (pJournalInquiryViewModel.PdfToDate != null)
            {
                mySql += " and AddedDate <='" + pJournalInquiryViewModel.PdfToDate.ToString("yyyy-MM-dd") + " 23:59:00' ";
            }
            if (!pJournalInquiryViewModel.PdfShowAll)
            {
                mySql += " and GLReferences.isPosted = 1";
            }

            //if (!string.IsNullOrEmpty(pJournalInquiryViewModel.AccountTypeICode))
            //{
            //    mySql += " and AccountTypeiCode ='" + pJournalInquiryViewModel.AccountTypeICode + "'";
            //}
            mySql += " order by VoucherNumber Desc";
            DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
            if (dtGLTransactions != null)
            {
                string pVal = "";
                foreach (DataRow MyRow in dtGLTransactions.Rows)
                {
                    pJournalInquiryViewModel.haveData = true;
                    strMyTable.Append("<tr>");
                    //if (MyRow[6].ToString() == Common.toString(pVal))
                    //{
                    //    strMyTable.Append("<td align='left' >&nbsp;</td>");
                    //    strMyTable.Append("<td align='left' >&nbsp;</td>");
                    //    strMyTable.Append("<td align='left' >&nbsp;</td>");
                    //}
                    //else
                    //{
                    strMyTable.Append("<td align='left'  >" + Common.toDateTime(MyRow["AddedDate"]).ToString("yyyy-MM-dd") + "</td>");
                    strMyTable.Append("<td align='left' >" + "<a href=\"javascript: showDetails('/Inquiries/GLTransView?Id=" + Security.EncryptQueryString(MyRow["GLReferenceID"].ToString()) + "','Voucher View')\")>" + MyRow["VoucherNumber"].ToString() + "</a></td>");
                    strMyTable.Append("<td align='left' >" + MyRow["Memo"].ToString() + "</td>");
                    //}
                    strMyTable.Append("<td align='left'>" + MyRow["AccountTitle"].ToString() + "</td>");
                    strMyTable.Append("<td align='left'>" + MyRow["ForeignCurrencyISOCode"].ToString() + "</td>");
                    strMyTable.Append("<td align='center'>" + DisplayUtils.GetSystemAmountFormat(Math.Abs(Common.toDecimal(MyRow["ForeignCurrencyAmount"]))) + "</td>");
                    decimal LCAmount = Common.toDecimal(MyRow["LocalCurrencyAmount"]);
                    LCAmount = DisplayUtils.getRoundDigit(Common.toString(MyRow["LocalCurrencyAmount"]));
                    if (LCAmount >= 0)
                    {
                        strMyTable.Append("<td align='left'>" + DisplayUtils.GetSystemAmountFormat(LCAmount) + "</td>");
                        strMyTable.Append("<td align='left'>&nbsp;</td>");
                    }
                    else
                    {
                        strMyTable.Append("<td align='left'>&nbsp;</td>");
                        strMyTable.Append("<td align='left'>" + DisplayUtils.GetSystemAmountFormat(Math.Abs(LCAmount)) + "</td>");
                    }

                    strMyTable.Append("<td align='left'>" + MyRow["PostedBy"].ToString() + "</td>");
                    strMyTable.Append("<td align='left'>" + MyRow["AuthBy"].ToString() + "</td>");
                    strMyTable.Append("<td align='left'>" + MyRow["TypeiCode"].ToString() + "</td>");
                    strMyTable.Append("</tr>");
                    pVal = Common.toString(MyRow[6]);
                }
            }
            strMyTable.Append("</tbody></table></div>");
            pJournalInquiryViewModel.Table = strMyTable.ToString();
            return pJournalInquiryViewModel;
        }

        //public void TransListExportToPDF(JournalInquiryViewModel pJournalInquiryViewModel)
        //{
        //    string HostName = System.Web.HttpContext.Current.Request.Url.Host;
        //    pJournalInquiryViewModel = JournalInquiryTable(pJournalInquiryViewModel);
        //    string fileName = Common.getRandomDigit(4) + "-" + "JournalInquiry";
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/");
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
        //    Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".pdf"));
        //    Response.End();
        //}
        #endregion
        public void TransListExportToExcel(JournalInquiryViewModel pJournalInquiryViewModel)
        {
            try
            {
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
                string mySql = "Select GLTransactions.GLReferenceID,  (Select AccountName From GLChartOfAccounts Where GLTransactions.AccountCode = GLChartOfAccounts.AccountCode) As AccountTitle, GLTransactions.AddedDate, GLTransactions.ForeignCurrencyAmount, GLTransactions.LocalCurrencyAmount, GLTransactions.Memo, GLReferences.VoucherNumber, GLTransactions.TransactionNumber As TransNo, GLTransactions.ForeignCurrencyISOCode, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.PostedBy) as PostedBy, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.AuthorizedBy) as AuthBy,(select Description from ConfigItemsData where GLReferences.TypeiCode= ConfigItemsData.ItemsDataCode) as TypeiCode FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID Where TransactionID > 0";
                //GLReferences.isPosted = 1 and GLTransactions.AccountCode = '" + myAccountCode + "'";
                if (!string.IsNullOrEmpty(pJournalInquiryViewModel.VoucherNumber))
                {
                    mySql += " and GLReferences.VoucherNumber='" + pJournalInquiryViewModel.VoucherNumber + "'";
                }
                if (!string.IsNullOrEmpty(pJournalInquiryViewModel.Memo))
                {
                    mySql += " and GLReferences.Memo  like '%" + pJournalInquiryViewModel.Memo + "%'";
                }
                if (pJournalInquiryViewModel.ExcelFromDate != null)
                {
                    mySql += " and AddedDate >='" + pJournalInquiryViewModel.ExcelFromDate.ToString("yyyy-MM-dd") + " 00:00:00' ";
                }
                if (pJournalInquiryViewModel.ExcelToDate != null)
                {
                    mySql += " and AddedDate <='" + pJournalInquiryViewModel.ExcelToDate.ToString("yyyy-MM-dd") + " 23:59:00' ";
                }
                if (!pJournalInquiryViewModel.ExcelShowAll)
                {
                    mySql += " and GLReferences.isPosted = 1";
                }

                //if (!string.IsNullOrEmpty(pJournalInquiryViewModel.AccountTypeICode))
                //{
                //    mySql += " and AccountTypeiCode ='" + pJournalInquiryViewModel.AccountTypeICode + "'";
                //}
                mySql += " order by VoucherNumber Desc";
                DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
                if (dtGLTransactions != null)
                {
                    string pVal = "";
                    foreach (DataRow MyRow in dtGLTransactions.Rows)
                    {
                        //if (MyRow[6].ToString() == Common.toString(pVal))
                        //{
                        //    worksheet.Cell("A" + mycount1).Value = "";
                        //    worksheet.Cell("B" + mycount1).Value = "";
                        //    worksheet.Cell("C" + mycount1).Value = "";
                        //}
                        //else
                        //{
                        worksheet.Cell("A" + mycount1).Value = Common.toDateTime(MyRow["AddedDate"]).ToString("yyyy-MM-dd");
                        worksheet.Cell("B" + mycount1).Value = MyRow["VoucherNumber"].ToString();
                        worksheet.Cell("C" + mycount1).Value = MyRow["Memo"].ToString();
                        // }
                        worksheet.Cell("D" + mycount1).Value = MyRow["AccountTitle"].ToString();
                        worksheet.Cell("E" + mycount1).Value = MyRow["ForeignCurrencyISOCode"].ToString();
                        worksheet.Cell("F" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(Math.Abs(Common.toDecimal(MyRow["ForeignCurrencyAmount"])));
                        decimal LCAmount = Common.toDecimal(MyRow["LocalCurrencyAmount"]);
                        LCAmount = DisplayUtils.getRoundDigit(Common.toString(MyRow["LocalCurrencyAmount"]));
                        if (LCAmount >= 0)
                        {
                            worksheet.Cell("G" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(LCAmount);
                            worksheet.Cell("H" + mycount1).Value = "";
                        }
                        else
                        {
                            worksheet.Cell("G" + mycount1).Value = "";
                            worksheet.Cell("H" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(LCAmount);
                        }
                        worksheet.Cell("I" + mycount1).Value = MyRow["PostedBy"].ToString();
                        worksheet.Cell("J" + mycount1).Value = MyRow["AuthBy"].ToString();
                        worksheet.Cell("K" + mycount1).Value = MyRow["TypeiCode"].ToString();
                        pVal = Common.toString(MyRow[6]);

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
                }

                string fileName = Common.getRandomDigit(4) + "-" + "JournalInquiry";
                workbook.SaveAs(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".xlsx"));
                Response.Clear();
                Response.ClearHeaders();
                Response.ClearContent();
                Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".xlsx");
                Response.ContentType = "text/plain";
                Response.Flush();
                //Response.TransmitFile(Server.MapPath("~/Uploads/" + "/1/DeanRpt.xlsx"));
                Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".xlsx"));
                Response.End();
            }
            catch (Exception x)
            {
                Response.Write("<font color='Red'>" + x.Message + " </font>");

            }
        }

        #endregion

        #region Trail Balance
        static string GetParentsString(string pAccountCode)
        {
            string path = "";
            if (pAccountCode.Trim() != "")
            {
                DataTable dtAccount = DBUtils.GetDataTable("select (select AccountName from GLChartOfAccounts b where  b.AccountCode=a.ParentAccountCode) as ParentAccountName, ParentAccountCode from  GLChartOfAccounts a where AccountCode='" + pAccountCode + "'");
                if (dtAccount != null)
                {
                    if (dtAccount.Rows.Count > 0)
                    {
                        path = Common.toString(dtAccount.Rows[0]["ParentAccountCode"]) + "," + Common.toString(dtAccount.Rows[0]["ParentAccountName"]);
                        //path = GetParentsString(Common.toString(dtAccount.Rows[0]["ParentAccountCode"])) + "->" + Common.toString(dtAccount.Rows[0]["AccountName"]) + path;
                    }
                }
            }
            return path;
        }

        public ActionResult TrailBalance(TrialbalanceViewModel pTrialbalanceViewModel)
        {
            pTrialbalanceViewModel.DateFrom = Common.toDateTime(SysPrefs.PostingDate).AddDays(-30); ;
            pTrialbalanceViewModel.DateTo = Common.toDateTime(SysPrefs.PostingDate);
            pTrialbalanceViewModel = TrailBalanceTableCopy(pTrialbalanceViewModel);
            return View("~/Views/Reports/Accounts/TrailBalance.cshtml", pTrialbalanceViewModel);
        }
        public ActionResult TrailBalance2(TrialbalanceViewModel pTrialbalanceViewModel)
        {
            pTrialbalanceViewModel = TrailBalanceTableCopy(pTrialbalanceViewModel);
            return View("~/Views/Reports/Accounts/TrailBalance.cshtml", pTrialbalanceViewModel);
        }
        public static TrialbalanceViewModel TrailBalanceTable(TrialbalanceViewModel pTrialbalanceViewModel)
        {
            pTrialbalanceViewModel.haveData = false;
            string mySql = "";
            StringBuilder strMyTable = new StringBuilder();
            strMyTable.Append("<b> Trail Balance </b>");
            strMyTable.Append(" as of: ");
            strMyTable.Append("" + @SysPrefs.PostingDate + "");
            strMyTable.Append("<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>");
            strMyTable.Append("<table style='table-layout: auto; empty-cells: show;' border='1' rules='rows'>");
            strMyTable.Append("<thead>");
            strMyTable.Append("<tr>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width: 15%;' align='center'>Main Classification</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width: 15%;' align='center'>GL Head</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' colspan='2' style='border: 1px solid antiquewhite; text-align:center;'>Total</th>");
            strMyTable.Append("</tr>");
            strMyTable.Append("<tr>");
            strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' scope='col' style='text-align:center; border: 1px solid antiquewhite;'>Debit</th>");
            strMyTable.Append("<th class='tabelHeader' scope='col' style='text-align:center; border: 1px solid antiquewhite;'>Credit</th>");
            strMyTable.Append("</tr>");
            strMyTable.Append("</thead>");
            strMyTable.Append("<tbody>");

            if (pTrialbalanceViewModel.ShowAll)
            {
                mySql = "SELECT GLChartOfAccounts.AccountCode, sum(GLTransactions.LocalCurrencyAmount) as LocalCurrencyBalance FROM GLChartOfAccounts LEFT OUTER JOIN GLTransactions ON GLChartOfAccounts.AccountCode = GLTransactions.AccountCode where isHead = 1 group by GLChartOfAccounts.AccountCode";
            }
            else
            {
                mySql = "SELECT GLChartOfAccounts.AccountCode, sum(GLTransactions.LocalCurrencyAmount) as LocalCurrencyBalance FROM GLChartOfAccounts INNER JOIN GLTransactions ON GLChartOfAccounts.AccountCode = GLTransactions.AccountCode where isHead = 1 group by GLChartOfAccounts.AccountCode";
            }
            //mySql = "SELECT GLChartOfAccounts.AccountName, GLChartOfAccounts.AccountCode, sum(GLTransactions.LocalCurrencyAmount) as LocalCurrencyBalance FROM GLChartOfAccounts INNER JOIN GLTransactions ON GLChartOfAccounts.AccountCode = GLTransactions.AccountCode where isHead = 1 group by GLChartOfAccounts.AccountCode, GLChartOfAccounts.AccountName";
            DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
            if (dtGLTransactions != null)
            {
                string subName1 = "";
                string subName2 = "";
                string subName3 = "";
                decimal TotalDebit = 0.0m;
                decimal TotalCredit = 0.0m;
                foreach (DataRow MyRow in dtGLTransactions.Rows)
                {
                    pTrialbalanceViewModel.haveData = true;
                    string AccountName = DBUtils.executeSqlGetSingle("Select AccountName from GLChartOfAccounts where AccountCode='" + MyRow["AccountCode"].ToString() + "' ");
                    strMyTable.Append("<tr>");
                    string subClass1 = GetParentsString(MyRow["AccountCode"].ToString());
                    string[] StrNameCode1 = subClass1.Split(',');
                    //if (Common.toString(StrNameCode1).Trim() != "")
                    if (StrNameCode1.Length > 0)
                    {
                        try
                        {
                            subName1 = StrNameCode1[1];
                            string subClass2 = GetParentsString(StrNameCode1[0]);
                            string[] StrNameCode2 = subClass2.Split(',');
                            //if (Common.toString(StrNameCode2).Trim() != "")
                            if (StrNameCode2.Length > 0)
                            {
                                subName2 = StrNameCode2[1];
                                string subClass3 = GetParentsString(StrNameCode2[0]);
                                // string[] StrNameCode3 = subClass2.Split(',');
                                string[] StrNameCode3 = subClass3.Split(',');

                                //if (Common.toString(StrNameCode3).Trim() != "")
                                if (StrNameCode3.Length > 0)
                                {
                                    subName3 = StrNameCode3[1];
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    strMyTable.Append("<td align='left' > " + subName3 + "</td>");
                    strMyTable.Append("<td align='left' > " + subName2 + "</td>");
                    strMyTable.Append("<td align='left' > " + subName1 + " </td>");
                    strMyTable.Append("<td align='left'  >" + AccountName + "</td>");
                    decimal localAmount = Common.toDecimal(MyRow["LocalCurrencyBalance"]);
                    if (localAmount < 0)
                    {
                        localAmount = Math.Abs(Common.toDecimal(localAmount));

                        strMyTable.Append("<td align='left' style='text-align:right;'> " + "-" + " </td>");
                        strMyTable.Append("<td align='left' style='text-align:right;'>" + "(" + localAmount + ")" + "</td>");
                        strMyTable.Append("<td align='left' style='text-align:right;'>" + "(" + localAmount + ")" + "</td>");
                        TotalCredit += localAmount;
                    }
                    else
                    {

                        strMyTable.Append("<td align='left' style='text-align:right;'>" + localAmount + "</td>");
                        strMyTable.Append("<td align='left' style='text-align:right;'> " + "-" + " </td>");
                        strMyTable.Append("<td align='left' style='text-align:right;'>" + localAmount + "</td>");
                        TotalDebit += localAmount;
                    }
                    //pVal = Common.toString(MyRow[6]);
                    strMyTable.Append("</tr>");
                }
                strMyTable.Append("<tr>");
                strMyTable.Append("<td colspan='5' style='background-color:rgb(255, 188, 0)'><b>Total</b></td>");
                strMyTable.Append("<td style='text-align:right; background-color:rgb(255, 188, 0)'>" + TotalDebit + "</td>");
                strMyTable.Append("<td style='text-align:right; background-color:rgb(255, 188, 0)'>" + "(" + TotalCredit + ")" + "</td>");
            }
            strMyTable.Append("</tbody></table></div>");
            pTrialbalanceViewModel.Table = strMyTable.ToString();
            return pTrialbalanceViewModel;
            //ltrTableBody.Text = Convert.ToString(strMyTable);
        }

        public static TrialbalanceViewModel TrailBalanceTableCopy(TrialbalanceViewModel pTrialbalanceViewModel)
        {
            PostingUtils.UpdateBalances(Common.toDateTime(@SysPrefs.PostingDate));
            pTrialbalanceViewModel.haveData = false;
            string mySql = "";
            StringBuilder strMyTable = new StringBuilder();
            strMyTable.Append("<b> Trail Balance </b>");
            strMyTable.Append(" From: ");
            strMyTable.Append("" + @SysPrefs.PostingDate + "");
            strMyTable.Append("<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>");
            strMyTable.Append("<table style='table-layout: auto; empty-cells: show;' border='1' rules='rows'>");
            strMyTable.Append("<thead>");
            strMyTable.Append("<tr>");
            strMyTable.Append("<th class='tabelHeader' colspan='2' style='border: 1px solid antiquewhite; width: 25%;' align='center'>Main Classification</th>");
            strMyTable.Append("<th class='tabelHeader' colspan='2' style='border: 1px solid antiquewhite; width: 25%;' align='center'>GL Head</th>");
            //strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' colspan='4' style='border: 1px solid antiquewhite; text-align:center;'>Total</th>");
            strMyTable.Append("</tr>");
            strMyTable.Append("<tr>");
            //strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' colspan='2' style='border: 1px solid antiquewhite; width: 25%;' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' colspan='2' style='border: 1px solid antiquewhite; width: 25%;' align='center'></th>");
            //strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            //strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' colspan='1' scope='col' style='text-align:center; border: 1px solid antiquewhite;'>Debit</th>");
            strMyTable.Append("<th class='tabelHeader' colspan='1' scope='col' style='text-align:center; border: 1px solid antiquewhite;'>Credit</th>");
            strMyTable.Append("<th class='tabelHeader' colspan='2' scope='col' style='text-align:center; border: 1px solid antiquewhite;'>Balance</th>");

            strMyTable.Append("</tr>");
            strMyTable.Append("</thead>");
            strMyTable.Append("<tbody>");

            //if (pTrialbalanceViewModel.ShowAll)
            //{
            mySql = "SELECT AccountName,AccountCode,sum(LocalCurrencyBalance) as LocalCurrencyBalance FROM GLChartOfAccounts where isHead = 1 group by AccountCode,AccountName order by AccountCode";
            //}
            if (pTrialbalanceViewModel.FromDate.ToShortDateString() != "01/01/0001" && pTrialbalanceViewModel.ToDate.ToShortDateString() != "01/01/0001")
            {
                mySql = "SELECT AccountName,AccountCode,sum(LocalCurrencyBalance) as LocalCurrencyBalance FROM GLChartOfAccounts where isHead = 1 and LocalCurrencyBalance!=0";
                mySql += $" and AddedDate <= CONVERT(datetime,'{pTrialbalanceViewModel.ToDate.ToShortDateString()}',103) and AddedDate >= CONVERT(datetime,'{pTrialbalanceViewModel.FromDate.ToShortDateString()}',103) ";
                mySql += "group by AccountCode,AccountName order by AccountCode";

            }
            DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
            if (dtGLTransactions != null)
            {
                string subName1 = "";
                string subName2 = "";
                string subName3 = "";
                decimal TotalDebit = 0.0m;
                decimal TotalCredit = 0.0m;
                decimal localAmount = 0;
                foreach (DataRow MyRow in dtGLTransactions.Rows)
                {
                    pTrialbalanceViewModel.haveData = true;
                    string AccountName = MyRow["AccountName"].ToString();
                    string AccountCode = MyRow["AccountCode"].ToString();
                    localAmount = Common.toDecimal(Math.Round(Common.toDecimal(MyRow["LocalCurrencyBalance"]), 2));
                    RecursiveAccountsViewModel pModel = PostingUtils.getParentAccounts(MyRow["AccountCode"].ToString());
                    if (pModel != null)
                    {
                        string[] StrNameCode1 = pModel.arr;
                        if (StrNameCode1.Length > 0)
                        {
                            for (int i = 0; i < StrNameCode1.Length; i++)
                            {
                                if (i == 0)
                                {
                                    subName3 = StrNameCode1[i].ToString();
                                }
                                else if (i == 1)
                                {
                                    subName2 = Common.toString(StrNameCode1[i]);
                                }
                                else if (i == 2)
                                {
                                    subName1 = Common.toString(StrNameCode1[i]);
                                }
                                else if (i == 3)
                                {
                                    if (Common.toString(StrNameCode1[3]).Trim() != AccountName.Trim())
                                    {
                                        subName1 += "=>" + Common.toString(StrNameCode1[i]);
                                    }
                                }
                            }
                        }
                    }
                    strMyTable.Append("<tr>");
                    strMyTable.Append("<td colspan='2' style='border: 1px solid antiquewhite; width: 25%;' align='left'> " + subName3 + " > " + subName2 + "</td>");
                    strMyTable.Append("<td colspan='2' style='border: 1px solid antiquewhite; width: 25%;' align='left'>" + AccountCode + " - " + AccountName + "</td>");
                    //decimal localAmount = Common.toDecimal(Math.Round(Common.toDecimal(MyRow["LocalCurrencyBalance"]), 2));

                    if (localAmount < 0)
                    {
                        localAmount = Math.Abs(Common.toDecimal(localAmount));

                        strMyTable.Append("<td colspan='1' align='left' style='text-align:center;'> " + "-" + " </td>");
                        strMyTable.Append("<td colspan='1' align='left' style='text-align:center;'>" + "(" + localAmount + ")" + "</td>");
                        strMyTable.Append("<td colspan='2' align='left' style='text-align:center;'>" + "(" + localAmount + ")" + "</td>");
                        TotalCredit += localAmount;
                    }
                    else
                    {
                        strMyTable.Append("<td colspan='1' align='left' style='text-align:center;'>" + localAmount + "</td>");
                        strMyTable.Append("<td colspan='1' align='left' style='text-align:center;'> " + "-" + " </td>");
                        strMyTable.Append("<td colspan='2' align='left' style='text-align:center;'>" + localAmount + "</td>");
                        TotalDebit += localAmount;
                    }
                    strMyTable.Append("</tr>");
                }
                strMyTable.Append("<tr>");
                strMyTable.Append("<td colspan='4' style='background-color:rgb(255, 188, 0)'><b>Total</b></td>");
                strMyTable.Append("<td colspan='1' style='text-align:center; background-color:rgb(255, 188, 0)'>" + TotalDebit + "</td>");
                strMyTable.Append("<td colspan='1' style='text-align:center; background-color:rgb(255, 188, 0)'>" + "(" + TotalCredit + ")" + "</td>");
                strMyTable.Append("<td colspan='2' style='text-align:center; background-color:rgb(255, 188, 0)'></td>");

                strMyTable.Append("</tr>");
            }
            strMyTable.Append("</tbody></table></div>");
            pTrialbalanceViewModel.Table = strMyTable.ToString();
            return pTrialbalanceViewModel;
            //ltrTableBody.Text = Convert.ToString(strMyTable);
        }
        public void ExportToExcel(TrialbalanceViewModel pTrialbalanceViewModel)
        {
            try
            {
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                decimal TotalDebit = 0.0m, TotalCredit = 0.0m;
                XLWorkbook workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Trail Balance Report");
                worksheet.Cell("A1").Value = "Trail Balance Report ";
                worksheet.Cell("A1").Style.Font.Bold = true;
                worksheet.Cell("A2").Value = "Trail Balance";
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("D2").Value = "As on Dated : " + @SysPrefs.PostingDate + "";
                worksheet.Cell("E2").Style.Font.Bold = true;

                worksheet.Cell("A3").Value = "Main Classification";
                worksheet.Cell("B3").Value = "Sub Classification1";
                worksheet.Cell("C3").Value = "Sub Classification2";
                worksheet.Cell("D3").Value = "Account Head";
                worksheet.Cell("E3").Value = "Balance";
                worksheet.Cell("F3").Value = "Debit";
                worksheet.Cell("G3").Value = "Credit";

                worksheet.Cell("A3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("B3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("C3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("D3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("E3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("F3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("G3").Style.Fill.BackgroundColor = XLColor.PersianRed;

                worksheet.Cell("A3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("B3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("C3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("D3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("E3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("F3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("G3").Style.Font.FontColor = XLColor.White;

                worksheet.ColumnWidth = 25;
                int mycount1 = 3;
                string mySql = "";

                if (pTrialbalanceViewModel.ShowAll)
                {
                    mySql = "SELECT AccountName,AccountCode,sum(LocalCurrencyBalance) as LocalCurrencyBalance FROM GLChartOfAccounts where isHead = 1 group by AccountCode,AccountName order by AccountCode";
                }
                else
                {
                    mySql = "SELECT AccountName,AccountCode,sum(LocalCurrencyBalance) as LocalCurrencyBalance FROM GLChartOfAccounts where isHead = 1 and LocalCurrencyBalance!=0 group by AccountCode,AccountName order by AccountCode";
                }
                DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
                if (dtGLTransactions != null)
                {
                    string subName1 = "", subName2 = "", subName3 = "", strDebit = "0", strCredit = "0", strAmount = "0";

                    foreach (DataRow MyRow in dtGLTransactions.Rows)
                    {
                        mycount1++;
                        string AccountName = MyRow["AccountName"].ToString();
                        RecursiveAccountsViewModel pModel = PostingUtils.getParentAccounts(MyRow["AccountCode"].ToString());
                        if (pModel != null)
                        {
                            string[] StrNameCode1 = pModel.arr;
                            if (StrNameCode1.Length > 0)
                            {
                                for (int i = 0; i < StrNameCode1.Length; i++)
                                {
                                    if (i == 0)
                                    {
                                        subName3 = StrNameCode1[i].ToString();
                                    }
                                    else if (i == 1)
                                    {
                                        subName2 = Common.toString(StrNameCode1[i]);
                                    }
                                    else if (i == 2)
                                    {
                                        subName1 = Common.toString(StrNameCode1[i]);
                                    }
                                    else if (i == 3)
                                    {
                                        if (Common.toString(StrNameCode1[3]).Trim() != AccountName.Trim())
                                        {
                                            subName1 += "=>" + Common.toString(StrNameCode1[i]);
                                        }
                                    }
                                }
                            }
                        }
                        decimal LocalCurrencyBalance = Common.toDecimal(MyRow["LocalCurrencyBalance"]);
                        if (LocalCurrencyBalance < 0)
                        {
                            LocalCurrencyBalance = Math.Abs(Common.toDecimal(LocalCurrencyBalance));
                            TotalCredit += LocalCurrencyBalance;
                            strCredit = "(" + LocalCurrencyBalance.ToString() + ")";
                            strDebit = "-";
                            strAmount = strCredit;
                        }
                        else
                        {
                            TotalDebit += LocalCurrencyBalance;
                            strCredit = "-";
                            strDebit = LocalCurrencyBalance.ToString();
                            strAmount = strDebit;
                        }
                        // table.Rows.Add(subName3, subName2, subName1, MyRow["AccountName"].ToString(), strAmount, strDebit, strCredit);
                        worksheet.Cell("A" + mycount1).Value = subName3;
                        worksheet.Cell("B" + mycount1).Value = subName2;
                        worksheet.Cell("C" + mycount1).Value = subName1;
                        worksheet.Cell("D" + mycount1).Value = MyRow["AccountName"].ToString();
                        worksheet.Cell("E" + mycount1).Value = strAmount;
                        worksheet.Cell("F" + mycount1).Value = strDebit;
                        worksheet.Cell("G" + mycount1).Value = strCredit;

                        worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("D" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("G" + mycount1).Style.Fill.BackgroundColor = XLColor.White;

                        worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("B" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("D" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("E" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("G" + mycount1).Style.Font.Bold = true;

                    }
                    mycount1++;
                    worksheet.Cell("A" + mycount1).Value = "Total";
                    worksheet.Cell("F" + mycount1).Value = TotalDebit;
                    worksheet.Cell("G" + mycount1).Value = "(" + TotalCredit + ")";
                    worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                    worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                    worksheet.Cell("G" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                    worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("G" + mycount1).Style.Font.Bold = true;
                }

                string fileName = Common.getRandomDigit(4) + "-" + "TrailBalanceReport";
                workbook.SaveAs(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".xlsx"));
                Response.Clear();
                Response.ClearHeaders();
                Response.ClearContent();
                Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".xlsx");
                Response.ContentType = "text/plain";
                Response.Flush();
                //Response.TransmitFile(Server.MapPath("~/Uploads/" + "/1/DeanRpt.xlsx"));
                Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".xlsx"));
                Response.End();

            }
            catch (Exception x)
            {
                Response.Write("<font color='Red'>" + x.Message + " </font>");
            }

        }

        //public void TrailExportToPDF(TrialbalanceViewModel pTrialbalanceViewModel)
        //{
        //    string HostName = System.Web.HttpContext.Current.Request.Url.Host;
        //    pTrialbalanceViewModel = TrailBalanceTableCopy(pTrialbalanceViewModel);
        //    string fileName = Common.getRandomDigit(4) + "-" + "TrailBalanceReport";
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/");
        //    HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //    htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //    PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(pTrialbalanceViewModel.Table, sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    Response.Clear();
        //    Response.ClearContent();
        //    Response.ClearHeaders();
        //    Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".pdf");
        //    Response.ContentType = "application/pdf";
        //    Response.Flush();
        //    Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".pdf"));
        //    Response.End();
        //}
        #endregion

        #region Translation Gain & Loss Report

        public ActionResult TranslationReport(TranslationGainLossModel pTranslationGainLossModel)
        {
            pTranslationGainLossModel = TranslationTable(pTranslationGainLossModel);
            return View("~/Views/Reports/Accounts/TranslationGainLossReport.cshtml", pTranslationGainLossModel);
        }

        public static TranslationGainLossModel TranslationTable(TranslationGainLossModel pTranslationGainLossModel)
        {
            PostingUtils.UpdateBalances(Common.toDateTime(@SysPrefs.PostingDate));
            decimal TotalRevisedAmount = 0.0m;
            decimal TotalProfitnLoss = 0.0m;
            decimal TotalLocalCurrencyBalance = 0.0m;
            string mySql = "";
            pTranslationGainLossModel.haveData = false;
            StringBuilder strMyTable = new StringBuilder();
            strMyTable.Append("<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>");
            strMyTable.Append("<table style='width:100%; border:0' cellspacing='0' cellpadding='5' align='center'>");
            strMyTable.Append("<tbody>");
            strMyTable.Append("<tr>");
            strMyTable.Append("<td width='500px'>");
            strMyTable.Append("<h4 class='media-heading' style='color: darkred; margin-left:5%'>" + @SysPrefs.SiteName + "</h4>");
            strMyTable.Append("</td>");
            strMyTable.Append("<td align='right'>");
            strMyTable.Append("Print Date Time : " + DateTime.Now + "");
            strMyTable.Append("</td></tr><tr><td width='300px'>");
            strMyTable.Append("<b>Revaluation Gain and Loss Report</b>");
            strMyTable.Append("</td><td>");
            strMyTable.Append("As on Dated : " + @SysPrefs.PostingDate + "");
            strMyTable.Append("</td>");
            strMyTable.Append("</tr>");
            strMyTable.Append("</tbody>");
            strMyTable.Append("</table>");
            strMyTable.Append("<table style='table-layout: auto; empty-cells: show; width:100%' border='1' rules='rows'>");
            strMyTable.Append("<thead>");
            strMyTable.Append("<tr>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite;' align='center'>Code</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite;' align='center'>Title of Account</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Curr</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Dr/Cr</th>");
            //strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Fc Balance</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Lc Balance</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Avg.Rate</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Rev.Rate</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Rev.Amount</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Profit/Loss</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Mode</th>");
            strMyTable.Append("</tr>");
            strMyTable.Append("</thead>");
            strMyTable.Append("<tbody>");
            // mySql = "Select GLChartOfAccounts.AccountCode, sum(GLTransactions.LocalCurrencyAmount) as LocalCurrencyBalance, sum(GLTransactions.ForeignCurrencyAmount) as ForeignCurrencyBalance  FROM GLChartOfAccounts INNER JOIN GLTransactions ON GLChartOfAccounts.AccountCode = GLTransactions.AccountCode where isHead = 1 and RevalueYN = 1 and CurrencyISOCode not in('GBP') group by GLChartOfAccounts.AccountCode";
            mySql = "Select AccountCode, AccountName, CurrencyISOCode, LocalCurrencyBalance FROM GLChartOfAccounts  where  LocalCurrencyBalance!=0 and isHead = 1 order by AccountCode";

            DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
            if (dtGLTransactions != null)
            {
                foreach (DataRow MyRow in dtGLTransactions.Rows)
                {
                    string DrCr = "";
                    pTranslationGainLossModel.haveData = true;
                    strMyTable.Append("<td align='left' > " + MyRow["AccountCode"].ToString() + "</td>");
                    strMyTable.Append("<td align='left' > " + MyRow["AccountName"].ToString() + "</td>");
                    strMyTable.Append("<td align='left' style='text-align:center;'> " + Common.toString(MyRow["CurrencyISOCode"]) + " </td>");
                    if (Common.toDecimal(MyRow["LocalCurrencyBalance"]) < 0)
                    {
                        DrCr = "Cr";
                        strMyTable.Append("<td align='left' style='text-align:center;' >" + "Cr" + "</td>");
                    }
                    else
                    {
                        DrCr = "Dr";
                        strMyTable.Append("<td align='left' style='text-align:center;' >" + "Dr" + "</td>");
                    }
                    // strMyTable.Append("<td align='left' style='text-align:right;'>" + Math.Round(Common.toDecimal(MyRow["ForeignCurrencyBalance"]), 2) + "</td>");
                    strMyTable.Append("<td align='left' style='text-align:right;'>" + Math.Round(Common.toDecimal(MyRow["LocalCurrencyBalance"]), 2) + "</td>");

                    //decimal ForeignCurrencyBalance = Common.toDecimal(MyRow["ForeignCurrencyBalance"]);
                    decimal LocalCurrencyBalance = Common.toDecimal(MyRow["LocalCurrencyBalance"]);
                    TotalLocalCurrencyBalance += LocalCurrencyBalance;
                    decimal AverageExchangeRate = 0.0m;
                    decimal CurrentExchangeRate = 0.0m;
                    //  string CurrenyType = Common.toString(MyRow["CurrenyType"]);
                    // CurrentExchangeRate = Common.toDecimal(MyRow["RevRate"]);
                    decimal RevisedAmount = 0.0m;
                    decimal ProfitnLoss = 0.0m;

                    // LocalCurrencyBalance = Math.Abs(LocalCurrencyBalance);
                    //  ForeignCurrencyBalance = Math.Abs(ForeignCurrencyBalance);
                    CurrentExchangeRate = Math.Round(CurrentExchangeRate, 2);
                    if (LocalCurrencyBalance != 0)
                    {
                        AverageExchangeRate = LocalCurrencyBalance;
                    }
                    AverageExchangeRate = Math.Round(AverageExchangeRate, 2);
                    if (CurrentExchangeRate != 0)
                    {
                        RevisedAmount = CurrentExchangeRate;
                    }
                    RevisedAmount = Math.Round(RevisedAmount, 2);

                    TotalRevisedAmount += RevisedAmount;
                    ProfitnLoss = RevisedAmount - LocalCurrencyBalance;
                    ProfitnLoss = Math.Round(ProfitnLoss, 2);
                    TotalProfitnLoss += ProfitnLoss;
                    //if (DrCr == "Cr")
                    //{
                    //    ProfitnLoss = LocalCurrencyBalance - RevisedAmount;
                    //    ProfitnLoss = Math.Round(ProfitnLoss, 2);
                    //    TotalProfitnLoss += ProfitnLoss;
                    //}
                    //else
                    //{
                    //    ProfitnLoss = RevisedAmount - LocalCurrencyBalance;
                    //    ProfitnLoss = Math.Round(ProfitnLoss, 2);
                    //    TotalProfitnLoss += ProfitnLoss;
                    //}

                    strMyTable.Append("<td align='left' style='text-align:right;'>" + AverageExchangeRate + "</td>");
                    strMyTable.Append("<td align='left' style='text-align:right;'>" + CurrentExchangeRate + "</td>");
                    strMyTable.Append("<td align='left' style='text-align:right;'>" + RevisedAmount + "</td>");
                    strMyTable.Append("<td align='left' style='text-align:right;'>" + ProfitnLoss + "</td>");
                    // strMyTable.Append("<td align='left' style='text-align:center;'>" + CurrenyType + "</td>");
                    strMyTable.Append("</tr>");


                }
                strMyTable.Append("<tr>");
                strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                strMyTable.Append("<td colspan='3' style='background-color:rgb(255, 188, 0)'><b>Total</b></td>");
                strMyTable.Append("<td align='left' style='text-align:right; background-color:rgb(255, 188, 0)'>" + Math.Round(TotalLocalCurrencyBalance, 2) + "</td>");
                strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                strMyTable.Append("<td align='left' style='text-align:right; background-color:rgb(255, 188, 0)'>" + TotalRevisedAmount + "</td>");
                strMyTable.Append("<td align='left' style='text-align:right; background-color:rgb(255, 188, 0)'>" + TotalProfitnLoss + "</td>");
                strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                strMyTable.Append("</tr>");
            }
            strMyTable.Append("</tbody></table></div>");
            pTranslationGainLossModel.Table = strMyTable.ToString();
            return pTranslationGainLossModel;
            //ltrTableBody.Text = Convert.ToString(strMyTable);
        }

        //public void TranslationExportToPrint(TranslationGainLossModel pTranslationGainLossModel)
        //{
        //    string HostName = System.Web.HttpContext.Current.Request.Url.Host;
        //    pTranslationGainLossModel = TranslationTable(pTranslationGainLossModel);
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp");
        //    HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //    htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //    //string fileName = Common.getRandomDigit(4) + "-" + "TranslationGainNLossReport";
        //    string fileName = "RevaluationGainNLossReportPDF_" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
        //    PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(pTranslationGainLossModel.Table, sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    Response.Flush();
        //    Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".pdf"));
        //    Response.End();

        //    //PdfPrinter 
        //    //PdfPrinter pdfPrinter = new PdfPrinter();

        //    //// optionally enable the silent printing
        //    //pdfPrinter.SilentPrinting = true;

        //    //// select the printer
        //    //pdfPrinter.PrinterSettings.PrinterName = "My Printer Name";

        //    //// send the PDF to printer
        //    //pdfPrinter.PrintPdf(pdfFile, fromPdfPageNumber, toPdfPageNumber);
        //}

        //public void TranslationExportToPDF(TranslationGainLossModel pTranslationGainLossModel)
        //{
        //    string HostName = System.Web.HttpContext.Current.Request.Url.Host;
        //    pTranslationGainLossModel = TranslationTable(pTranslationGainLossModel);
        //    //string fileName = Common.getRandomDigit(4) + "-" + "TranslateGainNLossReportPDF";
        //    string fileName = "RevaluationGainNLossReportPDF_" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp");
        //    HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //    htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //    PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(pTranslationGainLossModel.Table, sPathToWritePdfTo + "\\" + fileName + ".pdf");
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

        public ActionResult PostTranslationData(TranslationGainLossModel pTranslationGainLossModel)
        {


            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            ApplicationUser Profile = ApplicationUser.GetUserProfile();
            ////  string myConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
            bool isValid = true;
            string TypeiCode = "22JOURNAL";
            decimal TotalLocalCurrencyBalance = 0.0m;
            decimal TotalRevisedAmount = 0.0m;
            decimal TotalProfitnLoss = 0.0m;
            if (Common.ValidatePinCode(Profile.Id, pTranslationGainLossModel.PinCode))
            {
                if (Common.toString(SysPrefs.TranslationGainLossAccount).Trim() != "")
                {
                    //// using (SqlConnection connection = new SqlConnection(myConnectionString))
                    SqlConnection connection = Common.getConnection();
                    // {
                    SqlCommand InsertCommand = null;
                    StringBuilder strMyTable = new StringBuilder();

                    #region Html-Table-Header
                    strMyTable.Append("<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>");
                    strMyTable.Append("<table style='width:100%; border:0' cellspacing='0' cellpadding='5' align='center'>");
                    strMyTable.Append("<tbody>");
                    strMyTable.Append("<tr>");
                    strMyTable.Append("<td width='500px'>");
                    strMyTable.Append("<h4 class='media-heading' style='color: darkred; margin-left:5%'>" + @SysPrefs.SiteName + "</h4>");
                    strMyTable.Append("</td>");
                    strMyTable.Append("<td align='right'>");
                    strMyTable.Append("Print Date Time : " + DateTime.Now + "");
                    strMyTable.Append("</td></tr><tr><td width='300px'>");
                    strMyTable.Append("<b>Revaluation Gain and Loss Report</b>");
                    strMyTable.Append("</td><td>");
                    strMyTable.Append("As on Dated : " + @SysPrefs.PostingDate + "");
                    strMyTable.Append("</td>");
                    strMyTable.Append("</tr>");
                    strMyTable.Append("</tbody>");
                    strMyTable.Append("</table>");
                    strMyTable.Append("<table style='table-layout: auto; empty-cells: show; width:100%' border='1' rules='rows'>");
                    strMyTable.Append("<thead>");
                    strMyTable.Append("<tr>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite;' align='center'>Code</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite;' align='center'>Title of Account</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Curr</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Dr/Cr</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Fc Balance</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Lc Balance</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Avg.Rate</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Rev.Rate</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Rev.Amount</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Profit/Loss</th>");
                    strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; text-align:center;' align='center'>Mode</th>");
                    strMyTable.Append("</tr>");
                    strMyTable.Append("</thead>");
                    strMyTable.Append("<tbody>");
                    #endregion

                    XLWorkbook workbook = new XLWorkbook();

                    #region Worksheet-Header

                    var worksheet = workbook.Worksheets.Add("Revaluation Gain & Loss Report");
                    worksheet.Cell("A1").Value = @SysPrefs.SiteName;
                    worksheet.Cell("A1").Style.Font.Bold = true;
                    worksheet.Cell("A2").Value = "Revaluation Gain & Loss Report";
                    worksheet.Cell("A2").Style.Font.Bold = true;
                    worksheet.Cell("E2").Value = "As on Dated : " + @SysPrefs.PostingDate + "";
                    worksheet.Cell("E2").Style.Font.Bold = true;
                    worksheet.Cell("I1").Value = "Print Date Time : " + DateTime.Now + "";
                    worksheet.Cell("I1").Style.Font.Bold = true;

                    worksheet.Cell("A3").Value = "Account Code";
                    worksheet.Cell("B3").Value = "Account Name";
                    worksheet.Cell("C3").Value = "Curr";
                    worksheet.Cell("D3").Value = "DrCr";
                    worksheet.Cell("E3").Value = "F/C Balance";
                    worksheet.Cell("F3").Value = "L/C Balance";
                    worksheet.Cell("G3").Value = "Avg.Rate";
                    worksheet.Cell("H3").Value = "Current Rate";
                    worksheet.Cell("I3").Value = "Rev.Amount";
                    worksheet.Cell("J3").Value = "Profit/Loss";
                    worksheet.Cell("K3").Value = "Mode";

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

                    worksheet.Cell("A3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("B3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("C3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("D3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("E3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("F3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("G3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("H3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("I3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("J3").Style.Font.FontColor = XLColor.White;
                    worksheet.Cell("K3").Style.Font.FontColor = XLColor.White;
                    worksheet.ColumnWidth = 15;
                    #endregion

                    int mycount1 = 3;
                    try
                    {
                        Guid guid = Guid.NewGuid();
                        pTranslationGainLossModel.cpUserRefNo = guid.ToString();

                        string mySql = "";
                        string DrCr = "";
                        string AccountName = "";
                        string AccountCode = "";
                        string currencyISOCode = "";
                        decimal ForeignCurrencyBalance = 0.0m;
                        decimal LocalCurrencyBalance = 0.0m;
                        string sqlMyEntry = "";
                        decimal CurrentExchangeRate = 0.0m;

                        ////  mySql = "Select GLChartOfAccounts.AccountCode, sum(GLTransactions.LocalCurrencyAmount) as LocalCurrencyBalance, sum(GLTransactions.ForeignCurrencyAmount) as ForeignCurrencyBalance  FROM GLChartOfAccounts INNER JOIN GLTransactions ON GLChartOfAccounts.AccountCode = GLTransactions.AccountCode where isHead = 1 and RevalueYN = 1 and CurrencyISOCode not in('GBP') group by GLChartOfAccounts.AccountCode";
                        mySql = "Select AccountCode, AccountName, CurrencyISOCode,(Select CurrenyType from ListCurrencies where ListCurrencies.CurrencyISOCode = GLChartOfAccounts.CurrencyISOCode) as CurrenyType, (Select RevalueExchangeRate from ExchangeRates where AddedDate = '" + Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00") + "' and ExchangeRates.CurrencyISOCode = GLChartOfAccounts.CurrencyISOCode) as RevRate, LocalCurrencyBalance, ForeignCurrencyBalance FROM GLChartOfAccounts where (ForeignCurrencyBalance!=0 or LocalCurrencyBalance!=0) and isHead = 1 and RevalueYN = 1 and CurrencyISOCode not in('GBP') order by AccountCode";
                        DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
                        if (dtGLTransactions != null)
                        {
                            foreach (DataRow MyRow in dtGLTransactions.Rows)
                            {
                                mycount1++;
                                strMyTable.Append("<td align='left' > " + MyRow["AccountCode"].ToString() + "</td>");
                                strMyTable.Append("<td align='left' > " + MyRow["AccountName"].ToString() + "</td>");
                                strMyTable.Append("<td align='left' style='text-align:center;'> " + MyRow["CurrencyISOCode"].ToString() + " </td>");
                                if (Common.toDecimal(MyRow["LocalCurrencyBalance"]) < 0)
                                {
                                    DrCr = "Cr";
                                    strMyTable.Append("<td align='left' style='text-align:center;' >" + DrCr + "</td>");
                                }
                                else
                                {
                                    DrCr = "Dr";
                                    strMyTable.Append("<td align='left' style='text-align:center;' >" + DrCr + "</td>");
                                }
                                strMyTable.Append("<td align='left' style='text-align:right;'>" + Math.Round(Common.toDecimal(MyRow["ForeignCurrencyBalance"]), 2) + "</td>");
                                strMyTable.Append("<td align='left' style='text-align:right;'>" + Math.Round(Common.toDecimal(MyRow["LocalCurrencyBalance"]), 2) + "</td>");
                                AccountName = MyRow["AccountName"].ToString();
                                AccountCode = MyRow["AccountCode"].ToString();
                                currencyISOCode = MyRow["CurrencyISOCode"].ToString();
                                ForeignCurrencyBalance = Common.toDecimal(MyRow["ForeignCurrencyBalance"]);
                                LocalCurrencyBalance = Common.toDecimal(MyRow["LocalCurrencyBalance"]);
                                TotalLocalCurrencyBalance += LocalCurrencyBalance;
                                decimal AverageExchangeRate = 0.0m;
                                string CurrenyType = Common.toString(MyRow["CurrenyType"]);
                                CurrentExchangeRate = Common.toDecimal(MyRow["RevRate"]);
                                if (CurrentExchangeRate == 0)
                                {
                                    pTranslationGainLossModel.ErrMessage = Common.GetAlertMessage(1, "Incorrect exchange rate for currency " + currencyISOCode + ". Please udpate exchange rate and try again.");
                                    isValid = false;
                                    TotalProfitnLoss = 0;
                                }

                                decimal RevisedAmount = 0.0m;
                                decimal ProfitnLoss = 0.0m;
                                //if (LocalCurrencyBalance != 0 && ForeignCurrencyBalance != 0 && CurrentExchangeRate != 0)
                                //{

                                CurrentExchangeRate = Math.Round(CurrentExchangeRate, 2);
                                if (LocalCurrencyBalance != 0)
                                {
                                    AverageExchangeRate = ForeignCurrencyBalance / LocalCurrencyBalance;
                                }
                                AverageExchangeRate = Math.Round(AverageExchangeRate, 2);
                                if (CurrentExchangeRate != 0)
                                {
                                    RevisedAmount = ForeignCurrencyBalance / CurrentExchangeRate;
                                }
                                RevisedAmount = Math.Round(RevisedAmount, 2);

                                TotalRevisedAmount += RevisedAmount;
                                // ProfitnLoss = LocalCurrencyBalance - RevisedAmount;
                                ProfitnLoss = RevisedAmount - LocalCurrencyBalance;
                                ProfitnLoss = Math.Round(ProfitnLoss, 2);
                                TotalProfitnLoss += ProfitnLoss;
                                //LocalCurrencyBalance = Math.Abs(LocalCurrencyBalance);
                                //ForeignCurrencyBalance = Math.Abs(ForeignCurrencyBalance);
                                //AverageExchangeRate = ForeignCurrencyBalance / LocalCurrencyBalance;
                                //AverageExchangeRate = Math.Round(AverageExchangeRate, 4);
                                //RevisedAmount = ForeignCurrencyBalance / CurrentExchangeRate;
                                //RevisedAmount = Math.Round(RevisedAmount, 2);

                                //TotalRevisedAmount += RevisedAmount;
                                //if (DrCr == "Cr")
                                //{
                                //    ProfitnLoss = LocalCurrencyBalance - RevisedAmount;
                                //    ProfitnLoss = Math.Round(ProfitnLoss, 2);
                                //    TotalProfitnLoss += ProfitnLoss;
                                //}
                                //else
                                //{
                                //    ProfitnLoss = RevisedAmount - LocalCurrencyBalance;
                                //    ProfitnLoss = Math.Round(ProfitnLoss, 2);
                                //    TotalProfitnLoss += ProfitnLoss;
                                //}
                                //}

                                #region Worksheet & HtmlTable-Rows
                                worksheet.Cell("A" + mycount1).Value = MyRow["AccountCode"].ToString();
                                worksheet.Cell("B" + mycount1).Value = MyRow["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = MyRow["CurrencyISOCode"].ToString();
                                worksheet.Cell("D" + mycount1).Value = DrCr;
                                worksheet.Cell("E" + mycount1).Value = ForeignCurrencyBalance;
                                worksheet.Cell("F" + mycount1).Value = LocalCurrencyBalance;
                                worksheet.Cell("G" + mycount1).Value = AverageExchangeRate;
                                worksheet.Cell("H" + mycount1).Value = CurrentExchangeRate;
                                worksheet.Cell("I" + mycount1).Value = RevisedAmount;
                                worksheet.Cell("J" + mycount1).Value = ProfitnLoss;
                                worksheet.Cell("K" + mycount1).Value = CurrenyType;

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

                                strMyTable.Append("<td align='left' style='text-align:right;'>" + AverageExchangeRate + "</td>");
                                strMyTable.Append("<td align='left' style='text-align:right;'>" + CurrentExchangeRate + "</td>");
                                strMyTable.Append("<td align='left' style='text-align:right;'>" + RevisedAmount + "</td>");
                                strMyTable.Append("<td align='left' style='text-align:right;'>" + ProfitnLoss + "</td>");
                                strMyTable.Append("<td align='left' style='text-align:center;'>" + CurrenyType + "</td>");
                                strMyTable.Append("</tr>");

                                #endregion

                                #region inserting Gltemp Data

                                if (ProfitnLoss != 0)
                                {
                                    sqlMyEntry = "Insert into GLEntriesTemp (AccountCode,ForeignCurrencyAmount,LocalCurrencyAmount,Memo,GLTransactionTypeiCode,ExchangeRate, UserRefNo) Values (@AccountCode,@ForeignCurrencyAmount,@LocalCurrencyAmount,@Memo,@GLTransactionTypeiCode,@ExchangeRate,@UserRefNo)";

                                    InsertCommand = connection.CreateCommand();
                                    InsertCommand.Parameters.AddWithValue("@AccountCode", AccountCode);
                                    InsertCommand.Parameters.AddWithValue("@GLTransactionTypeiCode", TypeiCode);
                                    InsertCommand.Parameters.AddWithValue("@Memo", "Revaluation Gain & Loss Report");
                                    InsertCommand.Parameters.AddWithValue("@ForeignCurrencyAmount", 0);
                                    InsertCommand.Parameters.AddWithValue("@LocalCurrencyAmount", ProfitnLoss);
                                    InsertCommand.Parameters.AddWithValue("@ExchangeRate", CurrentExchangeRate);
                                    InsertCommand.Parameters.AddWithValue("@UserRefNo", pTranslationGainLossModel.cpUserRefNo);
                                    InsertCommand.CommandText = sqlMyEntry;
                                    InsertCommand.ExecuteNonQuery();
                                    InsertCommand.Parameters.Clear();
                                    //isValid = true;
                                }

                                #endregion
                            }
                            mycount1++;

                            #region Excel row
                            worksheet.Cell("A" + mycount1).Value = "Total";
                            worksheet.Cell("F" + mycount1).Value = TotalLocalCurrencyBalance;
                            worksheet.Cell("I" + mycount1).Value = TotalRevisedAmount;
                            worksheet.Cell("J" + mycount1).Value = TotalProfitnLoss;
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                            worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                            worksheet.Cell("I" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                            worksheet.Cell("J" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                            worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                            worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
                            worksheet.Cell("I" + mycount1).Style.Font.Bold = true;
                            worksheet.Cell("J" + mycount1).Style.Font.Bold = true;
                            #endregion

                            #region HtmlTable-row
                            strMyTable.Append("<tr>");
                            strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                            strMyTable.Append("<td colspan='3' style='background-color:rgb(255, 188, 0)'><b>Total</b></td>");
                            strMyTable.Append("<td align='left' style='text-align:right; background-color:rgb(255, 188, 0)'>" + Math.Round(TotalLocalCurrencyBalance, 2) + "</td>");
                            strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                            strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                            strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                            strMyTable.Append("<td align='left' style='text-align:right; background-color:rgb(255, 188, 0)'>" + TotalRevisedAmount + "</td>");
                            strMyTable.Append("<td align='left' style='text-align:right; background-color:rgb(255, 188, 0)'>" + TotalProfitnLoss + "</td>");
                            strMyTable.Append("<td style='background-color:rgb(255, 188, 0)'></td>");
                            strMyTable.Append("</tr>");
                            strMyTable.Append("</tbody></table></div>");
                            #endregion

                            if (TotalProfitnLoss != 0)
                            {
                                TotalProfitnLoss = -1 * TotalProfitnLoss;

                                sqlMyEntry = "Insert into GLEntriesTemp (AccountCode,ForeignCurrencyAmount,LocalCurrencyAmount,Memo,GLTransactionTypeiCode, ExchangeRate, UserRefNo, ForeignCurrencyISOCode) Values (@AccountCode,@ForeignCurrencyAmount,@LocalCurrencyAmount,@Memo,@GLTransactionTypeiCode,@ExchangeRate,@UserRefNo, @ForeignCurrencyISOCode)";
                                InsertCommand.Parameters.AddWithValue("@AccountCode", SysPrefs.TranslationGainLossAccount);
                                InsertCommand.Parameters.AddWithValue("@GLTransactionTypeiCode", TypeiCode);
                                InsertCommand.Parameters.AddWithValue("@Memo", "Revaluation Gain & Loss Report");
                                InsertCommand.Parameters.AddWithValue("@LocalCurrencyAmount", Math.Round(TotalProfitnLoss, 2));
                                InsertCommand.Parameters.AddWithValue("@ExchangeRate", CurrentExchangeRate);
                                InsertCommand.Parameters.AddWithValue("@ForeignCurrencyAmount", 0);
                                InsertCommand.Parameters.AddWithValue("@UserRefNo", pTranslationGainLossModel.cpUserRefNo);
                                InsertCommand.Parameters.AddWithValue("@ForeignCurrencyISOCode", currencyISOCode);
                                InsertCommand.CommandText = sqlMyEntry;
                                InsertCommand.ExecuteNonQuery();
                                InsertCommand.Parameters.Clear();
                                //isValid = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        //phMessage.Visible = true;
                        pTranslationGainLossModel.ErrMessage = pTranslationGainLossModel.ErrMessage + (pTranslationGainLossModel.ErrMessage == "" ? "" : "<br>Unable to Save Record Please try Again" + ex.ToString());
                        isValid = false;
                    }

                    #region Post for GL
                    if (isValid)
                    {
                        DataTable lstGLEntriesTemp = DBUtils.GetDataTable("select * from GLEntriesTemp where UserRefNo='" + pTranslationGainLossModel.cpUserRefNo + "'");
                        if (lstGLEntriesTemp != null)
                        {
                            if (lstGLEntriesTemp.Rows.Count > 0)
                            {
                                PostingViewModel PostingModel = PostingUtils.SaveGeneralJournelGLEntries(pTranslationGainLossModel.cpUserRefNo, "Translation Gain & Loss Report", TypeiCode, lstGLEntriesTemp, Profile.Id.ToString());
                                pTranslationGainLossModel.ErrMessage = PostingModel.ErrMessage;
                                if (PostingModel.isPosted)
                                {
                                    pTranslationGainLossModel.Table = strMyTable.ToString();
                                    ////string fileName = Common.getRandomDigit(4) + "-" + "TranslateGainNLossReportPDF";
                                    string fileName = "RevaluationGainNLossReportPDF_" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
                                    string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/");
                                    //HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
                                    //htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
                                    //PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(pTranslationGainLossModel.Table, sPathToWritePdfTo + "\\" + fileName + ".pdf");
                                    //pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
                                    string PdfFilePath = sPathToWritePdfTo + "\\" + fileName + ".pdf";

                                    ///fileName = Common.getRandomDigit(4) + "-" + "TranslationGainNLossReport";
                                    fileName = "RevaluationGainNLossReport_" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
                                    workbook.SaveAs(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".xlsx"));
                                    string ExcelFilePath = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + fileName + ".xlsx");

                                    string myPostingDate = Common.toString(SysPrefs.PostingDate);
                                    DateTime PostingDate = Common.toDateTime(myPostingDate);
                                    string strPostingDate = PostingDate.ToString("yyyy-MM-dd");

                                    string sqlMyEntry = "Insert into TranslationGainNLossLog(TotalAmount,GLReferenceID,PdfFileURL,ExcelFileURL,AddedBy, AddedDate) Values (@TotalAmount,@GLReferenceID,@PdfFileURL,@ExcelFileURL,@AddedBy, @AddedDate)";
                                    InsertCommand.Parameters.AddWithValue("@TotalAmount", TotalLocalCurrencyBalance);
                                    InsertCommand.Parameters.AddWithValue("@GLReferenceID", PostingModel.GLReferenceID);
                                    InsertCommand.Parameters.AddWithValue("@PdfFileURL", fileName + ".pdf");
                                    InsertCommand.Parameters.AddWithValue("@ExcelFileURL", fileName + ".xlsx");
                                    InsertCommand.Parameters.AddWithValue("@AddedBy", Profile.Id);
                                    InsertCommand.Parameters.AddWithValue("@AddedDate", strPostingDate);
                                    InsertCommand.CommandText = sqlMyEntry;
                                    InsertCommand.ExecuteNonQuery();
                                    InsertCommand.Parameters.Clear();
                                    isValid = true;
                                }

                            }
                            else
                            {
                                pTranslationGainLossModel.ErrMessage = Common.GetAlertMessage(1, "No temp data found.. cp:" + pTranslationGainLossModel.cpUserRefNo);
                            }
                        }
                        else
                        {
                            pTranslationGainLossModel.ErrMessage = Common.GetAlertMessage(1, "No temp data found. cp:" + pTranslationGainLossModel.cpUserRefNo);
                        }
                    }
                    else
                    {
                        if (Common.toString(pTranslationGainLossModel.ErrMessage).Trim() == "")
                        {
                            pTranslationGainLossModel.ErrMessage = Common.GetAlertMessage(1, "No valid data");
                        }
                    }
                    #endregion

                    connection.Close();
                    //// }
                }
                else
                {
                    pTranslationGainLossModel.ErrMessage = Common.GetAlertMessage(1, "Incorrect Translation Gain & Loss Account. Please use preferences to setup Translation Gain & Loss Account");

                }
                pTranslationGainLossModel = TranslationTable(pTranslationGainLossModel);
            }
            else
            {
                pTranslationGainLossModel.ErrMessage = Common.GetAlertMessage(1, "Invalid pin code entered.");
            }
            return View("~/Views/Reports/Accounts/TranslationGainLossReport.cshtml", pTranslationGainLossModel);
        }

        public ActionResult TranslationGainNLossLog(TranslationGainLossModel pTranslationGainLossModel)
        {
            string mySql = "";
            StringBuilder strMyTable = new StringBuilder();
            //strMyTable.Append("<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>");
            strMyTable.Append("<table style='table-layout: auto; empty-cells: show;' border='1' rules='rows'>");
            strMyTable.Append("<thead>");
            strMyTable.Append("<tr>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width:15%; text-align:center;' align='center'>Date</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width:10%; text-align:center;' align='center'>Voucher #</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width:10%; text-align:center;' align='center'>Amount</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width:15%; text-align:center;' align='center'>Added By</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width:15%; text-align:center;' align='center'>PDF File</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width:15%; text-align:center;' align='center'>Excel File</th>");
            strMyTable.Append("</tr>");
            strMyTable.Append("</thead>");
            strMyTable.Append("<tbody>");
            mySql = "Select (Select VoucherNumber from GLReferences where TranslationGainNLossLog.GLReferenceID=GLReferences.GLReferenceID) as VoucherNo,TotalAmount,PdfFileURL,ExcelFileURL, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= TranslationGainNLossLog.AddedBy) as AddedBy,AddedDate from TranslationGainNLossLog";
            DataTable dtLogData = DBUtils.GetDataTable(mySql);
            if (dtLogData != null)
            {
                foreach (DataRow MyRow in dtLogData.Rows)
                {
                    strMyTable.Append("<td align='center' > " + Common.toDateTime(MyRow["AddedDate"]).ToShortDateString() + " </td>");
                    strMyTable.Append("<td align='center' > " + MyRow["VoucherNo"].ToString() + "</td>");
                    strMyTable.Append("<td align='center' > " + MyRow["TotalAmount"].ToString() + "</td>");
                    strMyTable.Append("<td align='center' > " + MyRow["AddedBy"].ToString() + " </td>");
                    string PdfFile = MyRow["PdfFileURL"].ToString();
                    string ExcelFile = MyRow["ExcelFileURL"].ToString();
                    string[] ext = PdfFile.Split(new[] { "\\" }, StringSplitOptions.None);
                    string[] ext1 = ExcelFile.Split(new[] { "\\" }, StringSplitOptions.None);
                    string PdfFileName = ext.Last();
                    string ExcelFileName = ext1.Last();
                    strMyTable.Append("<td align='center' style='overflow-x:no-display'> " + "<a href='/RptAccounts/getFile?FileName=" + PdfFileName + "' target=''_blank''> " + "Click here" + " </a>" + "</td>");
                    strMyTable.Append("<td align='center' style='overflow-x:no-display'> " + "<a href='/RptAccounts/getFile?FileName=" + ExcelFileName + "' target=''_blank''> " + "Click here" + " </a> " + " </td>");
                    strMyTable.Append("</tr>");
                }
            }
            strMyTable.Append("</tbody></table>");
            pTranslationGainLossModel.Table = strMyTable.ToString();
            return View("~/Views/Reports/Accounts/TranslationGainNLossLog.cshtml", pTranslationGainLossModel);
        }
        public void getFile(string FileName)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            string extension = "";
            string[] ext = FileName.Split('.');
            extension = ext.Last();
            if (extension == "pdf")
            {
                Response.Clear();
                Response.ClearContent();
                Response.ClearHeaders();
                Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + FileName + ""));
                Response.AddHeader("Content-Disposition", "attachment; filename= " + FileName + ".pdf");
                Response.ContentType = "application/pdf";
                Response.Flush();
                Response.End();
            }
            else if (extension == "xlsx")
            {
                Response.Clear();
                Response.ClearContent();
                Response.ClearHeaders();
                Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp/" + FileName + ""));
                Response.AddHeader("Content-Disposition", "attachment; filename= " + FileName + ".xlsx");
                Response.ContentType = "application/xlsx";
                Response.Flush();
                Response.End();
            }
            else
            {
            }
        }

        public void TranslationExportToExcel(TranslationGainLossModel pTranslationGainLossModel)
        {
            string DrCr = "";
            try
            {
                XLWorkbook workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Translation Gain & Loss Report");
                worksheet.Cell("A1").Value = @SysPrefs.SiteName;
                worksheet.Cell("A1").Style.Font.Bold = true;
                worksheet.Cell("A2").Value = "Revaluation Gain & Loss Report";
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("E2").Value = "As on Dated : " + @SysPrefs.PostingDate + "";
                worksheet.Cell("E2").Style.Font.Bold = true;
                worksheet.Cell("I1").Value = "Print Date Time : " + DateTime.Now + "";
                worksheet.Cell("I1").Style.Font.Bold = true;

                worksheet.Cell("A3").Value = "Account Code";
                worksheet.Cell("B3").Value = "Account Name";
                worksheet.Cell("C3").Value = "Curr";
                worksheet.Cell("D3").Value = "DrCr";
                worksheet.Cell("E3").Value = "F/C Balance";
                worksheet.Cell("F3").Value = "L/C Balance";
                worksheet.Cell("G3").Value = "Avg.Rate";
                worksheet.Cell("H3").Value = "Current Rate";
                worksheet.Cell("I3").Value = "Rev.Amount";
                worksheet.Cell("J3").Value = "Profit/Loss";
                worksheet.Cell("K3").Value = "Mode";

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

                worksheet.Cell("A3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("B3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("C3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("D3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("E3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("F3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("G3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("H3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("I3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("J3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("K3").Style.Font.FontColor = XLColor.White;
                worksheet.ColumnWidth = 15;
                int mycount1 = 3;
                string mySql = "";
                decimal TotalLocalCurrencyBalance = 0.0m;
                decimal TotalRevisedAmount = 0.0m;
                decimal TotalProfitnLoss = 0.0m;
                //  mySql = "Select GLChartOfAccounts.AccountCode, sum(GLTransactions.LocalCurrencyAmount) as LocalCurrencyBalance, sum(GLTransactions.ForeignCurrencyAmount) as ForeignCurrencyBalance  FROM GLChartOfAccounts INNER JOIN GLTransactions ON GLChartOfAccounts.AccountCode = GLTransactions.AccountCode where isHead = 1 and RevalueYN = 1 and CurrencyISOCode not in('GBP') group by GLChartOfAccounts.AccountCode";
                mySql = "Select AccountCode, AccountName, CurrencyISOCode,(Select CurrenyType from ListCurrencies where ListCurrencies.CurrencyISOCode = GLChartOfAccounts.CurrencyISOCode) as CurrenyType, (Select RevalueExchangeRate from ExchangeRates where AddedDate = '" + Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00") + "' and ExchangeRates.CurrencyISOCode = GLChartOfAccounts.CurrencyISOCode) as RevRate, LocalCurrencyBalance, ForeignCurrencyBalance FROM GLChartOfAccounts where (ForeignCurrencyBalance!=0 or LocalCurrencyBalance!=0) and isHead = 1 and RevalueYN = 1 and CurrencyISOCode not in('GBP') order by AccountCode";

                DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
                if (dtGLTransactions != null)
                {
                    foreach (DataRow MyRow in dtGLTransactions.Rows)
                    {
                        mycount1++;
                        //  mySql = "Select AccountName,CurrencyISOCode, (Select CurrenyType from ListCurrencies where ListCurrencies.CurrencyISOCode = GLChartOfAccounts.CurrencyISOCode) as CurrenyType, (Select ExchangeRate from ExchangeRates where AddedDate = '" + Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00") + "' and ExchangeRates.CurrencyISOCode = GLChartOfAccounts.CurrencyISOCode) as RevRate from GLChartOfAccounts where AccountCode='" + MyRow["AccountCode"].ToString() + "'";
                        if (Common.toDecimal(MyRow["LocalCurrencyBalance"]) < 0)
                        {
                            DrCr = "Cr";
                        }
                        else
                        {
                            DrCr = "Dr";
                        }
                        decimal ForeignCurrencyBalance = Common.toDecimal(MyRow["ForeignCurrencyBalance"]);
                        decimal LocalCurrencyBalance = Common.toDecimal(MyRow["LocalCurrencyBalance"]);
                        TotalLocalCurrencyBalance += LocalCurrencyBalance;
                        decimal AverageExchangeRate = 0.0m;
                        decimal CurrentExchangeRate = 0.0m;
                        string CurrenyType = Common.toString(MyRow["CurrenyType"]);
                        CurrentExchangeRate = Common.toDecimal(MyRow["RevRate"]);
                        decimal RevisedAmount = 0.0m;
                        decimal ProfitnLoss = 0.0m;
                        //if (LocalCurrencyBalance != 0 && ForeignCurrencyBalance != 0)
                        // {
                        CurrentExchangeRate = Math.Round(CurrentExchangeRate, 2);
                        if (LocalCurrencyBalance != 0)
                        {
                            AverageExchangeRate = ForeignCurrencyBalance / LocalCurrencyBalance;
                        }
                        AverageExchangeRate = Math.Round(AverageExchangeRate, 2);
                        if (CurrentExchangeRate != 0)
                        {
                            RevisedAmount = ForeignCurrencyBalance / CurrentExchangeRate;
                        }
                        RevisedAmount = Math.Round(RevisedAmount, 2);

                        TotalRevisedAmount += RevisedAmount;
                        //  ProfitnLoss = LocalCurrencyBalance - RevisedAmount;
                        ProfitnLoss = RevisedAmount - LocalCurrencyBalance;
                        ProfitnLoss = Math.Round(ProfitnLoss, 2);
                        TotalProfitnLoss += ProfitnLoss;
                        //LocalCurrencyBalance = Math.Abs(LocalCurrencyBalance);
                        //ForeignCurrencyBalance = Math.Abs(ForeignCurrencyBalance);
                        //AverageExchangeRate = ForeignCurrencyBalance / LocalCurrencyBalance;
                        //AverageExchangeRate = Math.Round(AverageExchangeRate, 4);
                        //RevisedAmount = ForeignCurrencyBalance / CurrentExchangeRate;
                        //RevisedAmount = Math.Round(RevisedAmount, 2);

                        //TotalRevisedAmount += RevisedAmount;
                        //ProfitnLoss = LocalCurrencyBalance - RevisedAmount;
                        //ProfitnLoss = Math.Round(ProfitnLoss, 2);
                        //TotalProfitnLoss += ProfitnLoss;
                        //  }
                        worksheet.Cell("A" + mycount1).Value = MyRow["AccountCode"].ToString();
                        worksheet.Cell("B" + mycount1).Value = MyRow["AccountName"].ToString();
                        worksheet.Cell("C" + mycount1).Value = MyRow["CurrencyISOCode"].ToString();
                        worksheet.Cell("D" + mycount1).Value = DrCr;
                        worksheet.Cell("E" + mycount1).Value = ForeignCurrencyBalance;
                        worksheet.Cell("F" + mycount1).Value = LocalCurrencyBalance;
                        worksheet.Cell("G" + mycount1).Value = AverageExchangeRate;
                        worksheet.Cell("H" + mycount1).Value = CurrentExchangeRate;
                        worksheet.Cell("I" + mycount1).Value = RevisedAmount;
                        worksheet.Cell("J" + mycount1).Value = ProfitnLoss;
                        worksheet.Cell("K" + mycount1).Value = CurrenyType;

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
                    }
                    mycount1++;
                    worksheet.Cell("A" + mycount1).Value = "Total";
                    worksheet.Cell("E" + mycount1).Value = TotalLocalCurrencyBalance;
                    worksheet.Cell("I" + mycount1).Value = TotalRevisedAmount;
                    worksheet.Cell("J" + mycount1).Value = TotalProfitnLoss;

                    worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("D" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("I" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("J" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Cell("K" + mycount1).Style.Fill.BackgroundColor = XLColor.Yellow;

                    worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("I" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("J" + mycount1).Style.Font.Bold = true;

                }
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                //string fileName = Common.getRandomDigit(4) + "-" + "TranslationGainNLossReport";
                string fileName = "RevaluationGainNLossReportPDF-" + DateTime.Now.ToString("MMddyyyyHHmmssffff");
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

        #region Profit and Loss Income Statement
        public ActionResult AccountsGainAndLoss()
        {
            ChartOfAccountsModel pGainLossModel = new ChartOfAccountsModel();
            pGainLossModel.FromDate = Common.toDateTime(SysPrefs.PostingDate).AddDays(-30); ;
            pGainLossModel.ToDate = Common.toDateTime(SysPrefs.PostingDate);
            pGainLossModel.isvalid = false;
            pGainLossModel.ShowAll = false;
            return View("~/Views/Reports/Accounts/AccountsGainAndLoss.cshtml", pGainLossModel);
        }
        [HttpPost]
        public ActionResult AccountsGainAndLoss(ChartOfAccountsModel pGainLossModel)
        {
            pGainLossModel = BalanceGainAndLossTable(pGainLossModel);

            return View("~/Views/Reports/Accounts/AccountsGainAndLoss.cshtml", pGainLossModel);
        }

        public static ChartOfAccountsModel BuildIncomeTable(ChartOfAccountsModel pGainLossModel)
        {
            string RetVal = "";
            int mycount1 = 0; double pBalance = 0;
            pGainLossModel.TotalBalance = 0;
            pGainLossModel.tempTable = "";
            string sqlQuery = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + pGainLossModel.AccountCode + "'  UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
            DataTable dtAccount = DBUtils.GetDataTable(sqlQuery);
            if (dtAccount != null)
            {
                if (dtAccount.Rows.Count > 0)
                {
                    pGainLossModel.TotalBalance = GetBalance(Common.toString(pGainLossModel.AccountCode), pGainLossModel.FromDate, pGainLossModel.ToDate);
                }
                if (Common.toString(pGainLossModel.AccountCode).Trim() == Common.toString(SysPrefs.ProfileNLossIncomeAccount.Trim())) //For Income only
                {
                    //pGainLossModel.TotalBalance = Math.Abs(pGainLossModel.TotalBalance);
                    ////pGainLossModel.TotalBalance = pGainLossModel.TotalBalance;
                }

                foreach (DataRow drAccount in dtAccount.Rows)
                {
                    decimal level = 0; pBalance = 0;
                    decimal Currbalance = GetBalance(Common.toString(drAccount["AccountCode"]), pGainLossModel.FromDate, pGainLossModel.ToDate);
                    if (Common.toString(pGainLossModel.AccountCode).Trim() == Common.toString(SysPrefs.ProfileNLossIncomeAccount.Trim())) //For Income only
                    {
                        //// Currbalance = Math.Abs(Currbalance);
                    }
                    level = Common.toDecimal(drAccount["level"]);
                    bool isHead = Common.toBool(drAccount["isHead"]);
                    string pAccountName = drAccount["AccountName"].ToString();


                    //pGainLossModel.TotalBalance += Currbalance;
                    // pGainLossModel.TotalBalance = Math.Abs(pGainLossModel.TotalBalance);
                    if (Common.toDecimal(Currbalance) == 0)
                    {
                        if (isHead)
                        {
                            if (pGainLossModel.ShowAll)
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td>" + "&nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                pBalance = Common.toDouble(Currbalance);
                            }
                        }
                        else
                        {
                            RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                            RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                            RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                            pBalance = Common.toDouble(Currbalance);
                        }

                    }
                    else
                    {
                        if (Common.toDecimal(Currbalance) > 0)
                        {
                            if (isHead)
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td>" + "&nbsp; &nbsp; &nbsp;  " + drAccount["AccountName"] + "</td>";
                                RetVal += "<td></td><td align='right'>-" + Currbalance + "</td></tr>";
                                pBalance = -1 * Common.toDouble(Currbalance);
                            }
                            else
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                RetVal += "<td></td><td align='right'>-" + Currbalance + "</td></tr>";
                                pBalance = -1 * Common.toDouble(Currbalance);
                            }
                        }
                        else
                        {
                            if (isHead)
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td>" + "&nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                //  RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + " </td></tr>";
                                RetVal += "<td></td><td align='right'>" + Math.Abs(Currbalance) + " </td></tr>";
                                pBalance = Common.toDouble(Math.Abs(Currbalance));
                            }
                            else
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                // RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + "</td></tr>";
                                RetVal += "<td></td><td align='right'>" + Math.Abs(Currbalance) + " </td></tr>";
                                pBalance = Common.toDouble(Math.Abs(Currbalance));
                            }
                        }
                    }
                    pGainLossModel.TotalBalance += Common.toDecimal(pBalance);
                }
                pGainLossModel.tempTable = RetVal;
            }
            return pGainLossModel;
        }

        public static ChartOfAccountsModel BuildIncomeBalanceTable(ChartOfAccountsModel pGainLossModel)
        {
            string RetVal = "";
            int mycount1 = 0;
            pGainLossModel.TotalBalance = 0;
            pGainLossModel.tempTable = "";
            string sqlQuery = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + pGainLossModel.AccountCode + "'  UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
            DataTable dtAccount = DBUtils.GetDataTable(sqlQuery);
            if (dtAccount != null)
            {
                if (dtAccount.Rows.Count > 0)
                {
                    pGainLossModel.TotalBalance = GetBalance(Common.toString(pGainLossModel.AccountCode), pGainLossModel.FromDate, pGainLossModel.ToDate);
                }
                if (Common.toString(pGainLossModel.AccountCode).Trim() == Common.toString(SysPrefs.ProfileNLossIncomeAccount.Trim())) //For Income only
                {
                    pGainLossModel.TotalBalance = Math.Abs(pGainLossModel.TotalBalance);
                    ////pGainLossModel.TotalBalance = pGainLossModel.TotalBalance;
                }

                foreach (DataRow drAccount in dtAccount.Rows)
                {
                    decimal level = 0;
                    decimal Currbalance = GetBalance(Common.toString(drAccount["AccountCode"]), pGainLossModel.FromDate, pGainLossModel.ToDate);
                    if (Common.toString(pGainLossModel.AccountCode).Trim() == Common.toString(SysPrefs.ProfileNLossIncomeAccount.Trim())) //For Income only
                    {
                        //// Currbalance = Math.Abs(Currbalance);
                    }
                    level = Common.toDecimal(drAccount["level"]);
                    bool isHead = Common.toBool(drAccount["isHead"]);
                    string pAccountName = drAccount["AccountName"].ToString();
                    pGainLossModel.TotalBalance += Currbalance;
                    //pGainLossModel.TotalBalance = Math.Abs(pGainLossModel.TotalBalance);
                    if (Common.toDecimal(Currbalance) == 0)
                    {
                        if (isHead)
                        {
                            if (pGainLossModel.ShowAll)
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td>" + "&nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                            }
                        }
                        else
                        {
                            RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                            RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                            RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                        }

                    }
                    else
                    {
                        if (Common.toDecimal(Currbalance) > 0)
                        {
                            if (isHead)
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td>" + "&nbsp; &nbsp; &nbsp;  " + drAccount["AccountName"] + "</td>";
                                RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";

                            }
                            else
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                            }
                        }
                        else
                        {
                            if (isHead)
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td>" + "&nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                //  RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + " </td></tr>";
                                RetVal += "<td></td><td align='right'>" + Currbalance + " </td></tr>";
                            }
                            else
                            {
                                RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                // RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + "</td></tr>";
                                RetVal += "<td></td><td align='right'>" + Currbalance + " </td></tr>";
                            }
                        }
                    }
                }
                pGainLossModel.tempTable = RetVal;
            }
            return pGainLossModel;
        }

        public static ChartOfAccountsModel BalanceGainAndLossTable(ChartOfAccountsModel pGainLossModel)
        {
            pGainLossModel.isvalid = true;
            decimal TotalIncome = 0.0m;
            decimal TotalCostofSales = 0.0m;
            decimal TotalOtherIncome = 0.0m;
            decimal TotalAdministrativeCost = 0.0m;
            decimal TotalOtherExpenses = 0.0m;
            decimal GrossProfit = 0.0m;
            decimal NetProfit = 0.0m;

            #region Table Header
            string table = "";
            table += "<div class='row'>";
            table += "<div class='col-md-6'>";
            table += "<div class='media'>";
            table += "<div align='center' class='text-center' style='color:darkred;'><strong>" + @SysPrefs.SiteName + "</strong></div>";
            //table += "<div class='media -body'>";
            //table += "<h4 align='center' class='text-center'> <b>Profit And Loss</b></h4>";
            //table += "</div>";
            table += "</div>";
            table += "<br/><br/>";
            table += "<div class='media -body'>";
            table += "<h4 align='center' class='text-center'> <b>Profit And Loss</b></h4>";
            table += "</div>";
            table += "<br/><br/>";
            table += "<h4><b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; From Date : &nbsp; &nbsp;" + pGainLossModel.FromDate.ToShortDateString() + "&nbsp; &nbsp; ToDate: &nbsp; &nbsp;" + pGainLossModel.ToDate.ToShortDateString() + "</b></h4>";
            table += "</div>";
            table += "</div>";
            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' width='50%' cellpadding='5'>";
            }
            table += "<thead>";
            table += " <tr>";
            table += " </tr>";
            table += " <tr>";
            table += "<th style='background-color:#1c3a70; color:#FFF;'></th>";
            table += "<th style='background-color:#1c3a70; color:#FFF; text-align:right;'></th>";
            table += "<th style='background-color:#1c3a70; color:#FFF; text-align:right;'>GBP</th>";
            table += "</tr>";
            table += "<thead>";
            table += "<tbody>";
            table += "<tr style='border-bottom:#000 1px solid;'>";
            table += "<td><h5><b>Income</b><h5></td><td></td><td></td></tr>";

            #endregion
            //pGainLossModel.FromDate = PostingUtils.BeginFiscalYear();
            //pGainLossModel.ToDate = pGainLossModel.ToDate;            

            DateTime FromDate = Common.toDateTime(pGainLossModel.FromDate.ToString("MM/dd/yyyy"));
            DateTime ToDate = Common.toDateTime(pGainLossModel.ToDate.ToString("MM/dd/yyyy"));

            #region Income

            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' id='tblIncome' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' id='tblIncome' width='50%' cellpadding='5'>";
            }
            pGainLossModel.AccountCode = SysPrefs.ProfileNLossIncomeAccount;
            pGainLossModel = BuildIncomeTable(pGainLossModel);
            //pGainLossModel = BuildIncomeBalanceTable(pGainLossModel);
            table += pGainLossModel.tempTable;
            TotalIncome = pGainLossModel.TotalBalance;
            table += "</table>";
            table += "<table border='0' width='50%' cellpadding='5'>";
            //table += BuildBSTable(SysPrefs.BSNonCurrentAssetsAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            table += "<tr><td><b>Total income</b></td>";
            table += "<td></td><td align='right'><b>" + TotalIncome + "</b></td></tr>";
            //table += "<td></td><td align='right'><b>" + TotalTangibleAssets + "</b></td></tr>";

            table += "<tr><td height='30'></td></tr>";
            table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Cost Of Sales</b></h5></td><td></td><td></td></tr>";
            //table += "<tr><td><b> Cash at bank and in hand </b></td>";
            //table += "<td></td><td></td></tr>";
            table += "</table>";
            #endregion

            #region Cost of sales

            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' id='tblCostofSales' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' id='tblCostofSales' width='50%' cellpadding='5'>";
            }
            pGainLossModel.AccountCode = SysPrefs.ProfileNLossCostOfSales;
            pGainLossModel = BuildIncomeBalanceTable(pGainLossModel);
            //// pGainLossModel= BuildIncomeTable(pGainLossModel);
            table += pGainLossModel.tempTable;
            TotalCostofSales = pGainLossModel.TotalBalance;
            table += "</table>";

            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' width='50%' cellpadding='5'>";
            }
            //table += BuildBSTable(SysPrefs.BSCurrentAssetsAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            table += "<tr><td><b>Total cost of sales</b></td>";
            table += "<td></td><td align='right'><b>" + TotalCostofSales + "</b></td></tr>";

            table += "<tr><td height='30'></td></tr>";

            GrossProfit = TotalIncome - TotalCostofSales;
            table += "<tr style='border-bottom:#000 1px solid; border-top:#000 1px solid;'><td><b>Gross Profit</b></td><td></td>";
            table += "<td align='right'><b>" + GrossProfit + "</b></td></tr>";
            table += "<tr><td height='60'></td></tr>";
            table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Other income</b></h5></td><td></td><td></td></tr>";
            table += "</table>";
            #endregion

            #region Other income

            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' id='tblOtherincome' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' id='tblOtherincome' width='50%' cellpadding='5'>";
            }
            pGainLossModel.AccountCode = SysPrefs.ProfileNLossOtherIncome;
            pGainLossModel = BuildIncomeTable(pGainLossModel);
            table += pGainLossModel.tempTable;
            TotalOtherIncome = pGainLossModel.TotalBalance;
            table += "</table>";
            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' width='50%' cellpadding='5'>";
            }
            //table += BuildBSTable(SysPrefs.BSCurrentLiabilitesAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            table += "<tr><td><b>Total other income</b></td>";
            if (TotalOtherIncome > 0)
            {
                table += "<td></td><td align='right'><b>" + TotalOtherIncome + "</b></td></tr>";
            }
            else
            {
                // table += "<td></td><td align='right'><b>(" + Math.Abs(TotalOtherIncome) + ")</b></td></tr>";
                table += "<td></td><td align='right'><b>" + Math.Abs(TotalOtherIncome) + "</b></td></tr>";
            }

            table += "<tr><td height='30'></td></tr>";
            table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Administrative Costs</b></h5></td><td></td><td></td></tr>";
            table += "</table>";
            #endregion

            #region Administrative costs

            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' id='tblAdministrativeCosts' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' id='tblAdministrativeCosts' width='50%' cellpadding='5'>";
            }
            pGainLossModel.AccountCode = SysPrefs.ProfileNLossAdministrativeExpenses;
            pGainLossModel = BuildIncomeBalanceTable(pGainLossModel);
            table += pGainLossModel.tempTable;
            TotalAdministrativeCost = pGainLossModel.TotalBalance;
            table += "</table>";
            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' width='50%' cellpadding='5'>";
            }
            //table += BuildBSTable(SysPrefs.BSNonCurrentliabilitesAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            table += "<tr><td><b>Total administrative costs</b></td>";
            table += "<td></td><td align='right'><b>" + TotalAdministrativeCost + "</b></td></tr>";
            table += "<tr><td height='60'></td></tr>";
            table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Other Expense</b></h5></td><td></td><td></td></tr>";
            table += "</table>";
            #endregion

            #region  Other Expense

            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' id='tblOtherExpense' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' id='tblOtherExpense' width='50%' cellpadding='5'>";
            }
            pGainLossModel.AccountCode = SysPrefs.ProfileNLossOtherExpenses;
            pGainLossModel = BuildIncomeBalanceTable(pGainLossModel);
            table += pGainLossModel.tempTable;
            TotalOtherExpenses = pGainLossModel.TotalBalance;
            table += "</table>";
            if (pGainLossModel.isPrint)
            {
                table += "<table border='0' width='90%' cellpadding='5'>";
            }
            else
            {
                table += "<table border='0' width='50%' cellpadding='5'>";
            }
            //table += BuildBSTable(SysPrefs.BSCapitalNReservesAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            table += "<tr><td><b>Total other expense</b></td>";
            table += "<td></td><td align='right'><b>" + Math.Abs(TotalOtherExpenses) + "</b></td></tr>";

            table += "<tr><td height='30'></td></tr>";
            decimal NetProfitA = GrossProfit + Math.Abs(TotalOtherIncome);
            decimal NetProfitB = NetProfitA - TotalAdministrativeCost;
            NetProfit = NetProfitB - Math.Abs(TotalOtherExpenses);
            table += "<tr style='border-bottom:#000 1px solid; border-top:#000 1px solid;'><td><b>Net Profit</b></td><td></td>";
            table += "<td align='right'><b>" + NetProfit + "</b></td></tr>";
            table += "</table>";
            #endregion

            table += "</tbody>";
            table += "</table>";
            pGainLossModel.HtmlTable = table;
            return pGainLossModel;
        }

        //public static ChartOfAccountsModel BalanceGainAndLossTable(ChartOfAccountsModel pGainLossModel)
        //{
        //    pGainLossModel.isvalid = true;
        //    decimal Balance1 = 0.0m;
        //    decimal Balance2 = 0.0m;
        //    decimal Balance3 = 0.0m;
        //    decimal Balance4 = 0.0m;
        //    decimal Balance5 = 0.0m;
        //    decimal TotalIncome = 0.0m;
        //    decimal TotalCostofSales = 0.0m;
        //    decimal TotalOtherIncome = 0.0m;
        //    decimal TotalAdministrativeCost = 0.0m;
        //    decimal TotalOtherExpenses = 0.0m;
        //    decimal GrossProfit = 0.0m;
        //    decimal NetProfit = 0.0m;
        //    string table = "";
        //    table += "<div class='row'>";
        //    table += "<div class='col-md-6'>";
        //    table += "<div class='media'>";
        //    table += "<div class='media -body'>";
        //    table += "<h4 align='center' class='text-center'><b>Profit And Loss</b></h4>";
        //    table += "</div>";
        //    table += "</div>";
        //    table += "<br/><br/>";
        //    table += "<div align='center' class='text-center' style='color:darkred;'><strong>" + @SysPrefs.SiteName + "</strong></div>";
        //    table += "<br/><br/>";
        //    table += "<h4><b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; From:&nbsp; &nbsp;" + pGainLossModel.FromDate.ToShortDateString() + "&nbsp; &nbsp; &nbsp; &nbsp; To: &nbsp; &nbsp;" + pGainLossModel.ToDate.ToShortDateString() + "</b></h4>";
        //    table += "</div>";
        //    table += "</div>";
        //    table += "<table border='0' width='50%' cellpadding='5'>";
        //    table += "<thead>";
        //    table += " <tr>";
        //    table += " </tr>";
        //    table += " <tr>";
        //    table += "<th style='background-color:#1c3a70; color:#FFF;'></th>";
        //    table += "<th style='background-color:#1c3a70; color:#FFF; text-align:right;'></th>";
        //    table += "<th style='background-color:#1c3a70; color:#FFF; text-align:right;'>GBP</th>";
        //    table += "</tr>";
        //    table += "<thead>";
        //    table += "<tbody>";
        //    table += "<tr style='border-bottom:#000 1px solid;'>";
        //    table += "<td><h5><b>Income</b><h5></td><td></td><td></td></tr>";

        //    string SqlIncome = "; WITH ret AS(SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossIncomeAccount + "' UNION ALL SELECT t.AccountId, t.AccountCode, t.AccountName, t.ParentAccountCode, t.isHead FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead, (select isnull(sum(LocalCurrencyAmount), 0) from GLTransactions where GLTransactions.AccountCode = ret.AccountCode and GLTransactions.AddedDate >= '" + pGainLossModel.FromDate.ToString("MM/dd/yyyy") + "' and GLTransactions.AddedDate <= '" + pGainLossModel.ToDate.ToString("MM/dd/yyyy") + "') as bal FROM ret where isHead = 1";
        //    DataTable dtIncome = DBUtils.GetDataTable(SqlIncome);
        //    if (dtIncome != null && dtIncome.Rows.Count > 0)
        //    {
        //        foreach (DataRow dr in dtIncome.Rows)
        //        {
        //            if (Common.toDecimal(dr["bal"]) == 0)
        //            {
        //                if (pGainLossModel.ShowAll)
        //                {
        //                    table += "<tr><td>" + dr["AccountName"].ToString() + "</td>";
        //                    Balance1 = Math.Round(Common.toDecimal(dr["bal"]), 2);
        //                    TotalIncome += Balance1;
        //                    table += "<td></td><td align='right'>" + Balance1 + "</td></tr>";
        //                }
        //            }
        //            else
        //            {
        //                if (Common.toDecimal(dr["bal"]) > 0)
        //                {
        //                    table += "<tr><td>" + dr["AccountName"].ToString() + "</td>";
        //                    Balance1 = Math.Round(Common.toDecimal(dr["bal"]) * (-1), 2);
        //                    TotalIncome += Balance1;
        //                    table += "<td></td><td align='right'>" + Balance1 + "</td></tr>";
        //                }
        //                else
        //                {
        //                    table += "<tr><td>" + dr["AccountName"].ToString() + "</td>";
        //                    Balance1 = Math.Abs(Math.Round(Common.toDecimal(dr["bal"]), 2));
        //                    TotalIncome += Balance1;
        //                    table += "<td></td><td align='right'>" + Balance1 + "</td></tr>";
        //                }
        //            }
        //        }
        //    }
        //    table += "<tr><td><b>Total income</b></td>";
        //    table += "<td></td><td align='right'><b>" + TotalIncome + "</b></td></tr>";
        //    table += "<tr><td height='30'></td></tr>";


        //    table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Cost of sales</b></h5></td><td></td><td></td></tr>";

        //    string SqlCostofSales = "; WITH ret AS(SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossCostOfSales + "' UNION ALL SELECT t.AccountId, t.AccountCode, t.AccountName, t.ParentAccountCode, t.isHead FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead, (select isnull(sum(LocalCurrencyAmount), 0) from GLTransactions where GLTransactions.AccountCode = ret.AccountCode and GLTransactions.AddedDate >= '" + pGainLossModel.FromDate.ToString("MM/dd/yyyy") + "' and GLTransactions.AddedDate <= '" + pGainLossModel.ToDate.ToString("MM/dd/yyyy") + "') as bal FROM ret where isHead = 1";
        //    DataTable dtCostofSales = DBUtils.GetDataTable(SqlCostofSales);
        //    if (dtCostofSales != null && dtCostofSales.Rows.Count > 0)
        //    {
        //        foreach (DataRow dr1 in dtCostofSales.Rows)
        //        {
        //            if (Common.toDecimal(dr1["bal"]) == 0)
        //            {
        //                if (pGainLossModel.ShowAll)
        //                {
        //                    table += "<tr><td>" + dr1["AccountName"].ToString() + "</td>";
        //                    Balance2 = Math.Round(Common.toDecimal(dr1["bal"]), 2);
        //                    TotalCostofSales += Balance2;
        //                    table += "<td></td><td align='right'>" + Balance2 + "</td></tr>";
        //                }
        //            }
        //            else
        //            {
        //                table += "<tr><td>" + dr1["AccountName"].ToString() + "</td>";
        //                Balance2 = Math.Round(Common.toDecimal(dr1["bal"]), 2);
        //                TotalCostofSales += Balance2;
        //                table += "<td></td><td align='right'>" + Balance2 + "</td></tr>";
        //            }
        //        }
        //    }
        //    table += "<tr><td><b>Total Cost of sales</b></td>";
        //    table += "<td></td><td align='right'><b>" + TotalCostofSales + "</b></td></tr>";
        //    table += "<tr><td height='30'></td></tr>";

        //    GrossProfit = TotalIncome - TotalCostofSales;
        //    table += "<tr style='border-bottom:#000 1px solid; border-top:#000 1px solid;'><td><b>Gross Profit</b></td><td></td>";
        //    table += "<td align='right'><b>" + GrossProfit + "</b></td></tr>";
        //    table += "<tr><td height='60'></td></tr>";
        //    table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Other income</b></h5></td><td></td><td></td></tr>";

        //    string SqlOtherIncome = "; WITH ret AS(SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossOtherIncome + "' UNION ALL SELECT t.AccountId, t.AccountCode, t.AccountName, t.ParentAccountCode, t.isHead FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead, (select isnull(sum(LocalCurrencyAmount), 0) from GLTransactions where GLTransactions.AccountCode = ret.AccountCode and GLTransactions.AddedDate >= '" + pGainLossModel.FromDate.ToString("MM/dd/yyyy") + "' and GLTransactions.AddedDate <= '" + pGainLossModel.ToDate.ToString("MM/dd/yyyy") + "') as bal FROM ret where isHead = 1";
        //    DataTable dtOtherIncome = DBUtils.GetDataTable(SqlOtherIncome);
        //    if (dtOtherIncome != null && dtOtherIncome.Rows.Count > 0)
        //    {
        //        foreach (DataRow dr2 in dtOtherIncome.Rows)
        //        {
        //            if (Common.toDecimal(dr2["bal"]) == 0)
        //            {
        //                if (pGainLossModel.ShowAll)
        //                {
        //                    table += "<tr><td>" + dr2["AccountName"].ToString() + "</td>";
        //                    Balance3 = Math.Round(Common.toDecimal(dr2["bal"]), 2);
        //                    TotalOtherIncome += Balance3;
        //                    table += "<td></td><td align='right'>" + Balance3 + "</td></tr>";
        //                }
        //            }
        //            else
        //            {
        //                if (Common.toDecimal(dr2["bal"]) > 0)
        //                {
        //                    table += "<tr><td>" + dr2["AccountName"].ToString() + "</td>";
        //                    Balance3 = Math.Round(Common.toDecimal(dr2["bal"]) * (-1), 2);
        //                    TotalOtherIncome += Balance3;
        //                    table += "<td></td><td align='right'>" + Balance3 + "</td></tr>";
        //                }
        //                else
        //                {
        //                    table += "<tr><td>" + dr2["AccountName"].ToString() + "</td>";
        //                    Balance3 = Math.Abs(Math.Round(Common.toDecimal(dr2["bal"]), 2));
        //                    TotalOtherIncome += Balance3;
        //                    table += "<td></td><td align='right'>" + Balance3 + "</td></tr>";
        //                }
        //            }
        //        }
        //    }
        //    table += "<tr><td><b>Total Other income</b></td>";
        //    table += "<td></td><td align='right'><b>" + TotalOtherIncome + "</b></td></tr>";
        //    table += "<tr><td height='30'></td></tr>";
        //    table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Administrative costs</b></h5></td><td></td><td></td></tr>";

        //    string SqlAdministrativeCost = "; WITH ret AS(SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossAdministrativeExpenses + "' UNION ALL SELECT t.AccountId, t.AccountCode, t.AccountName, t.ParentAccountCode, t.isHead FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead, (select isnull(sum(LocalCurrencyAmount), 0) from GLTransactions where GLTransactions.AccountCode = ret.AccountCode and GLTransactions.AddedDate >= '" + pGainLossModel.FromDate.ToString("MM/dd/yyyy") + "' and GLTransactions.AddedDate <= '" + pGainLossModel.ToDate.ToString("MM/dd/yyyy") + "') as bal FROM ret where isHead = 1";
        //    DataTable dtAdministrativeCost = DBUtils.GetDataTable(SqlAdministrativeCost);
        //    if (dtAdministrativeCost != null && dtAdministrativeCost.Rows.Count > 0)
        //    {
        //        foreach (DataRow dr3 in dtAdministrativeCost.Rows)
        //        {
        //            if (Common.toDecimal(dr3["bal"]) == 0)
        //            {
        //                if (pGainLossModel.ShowAll)
        //                {
        //                    table += "<tr><td>" + dr3["AccountName"].ToString() + "</td>";
        //                    Balance4 = Math.Round(Common.toDecimal(dr3["bal"]), 2);
        //                    TotalAdministrativeCost += Balance4;
        //                    table += "<td></td><td align='right'>" + Balance4 + "</td></tr>";
        //                }
        //            }
        //            else
        //            {
        //                table += "<tr><td>" + dr3["AccountName"].ToString() + "</td>";
        //                Balance4 = Math.Round(Common.toDecimal(dr3["bal"]), 2);
        //                TotalAdministrativeCost += Balance4;
        //                table += "<td></td><td align='right'>" + Balance4 + "</td></tr>";
        //            }

        //        }
        //    }
        //    table += "<tr><td><b>Total Administrative costs</b></td>";
        //    table += "<td></td><td align='right'><b>" + TotalAdministrativeCost + "</b></td></tr>";
        //    table += "<tr><td height='30'></td></tr>";

        //    string SqlOtherExpenses = "; WITH ret AS(SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossOtherExpenses + "' UNION ALL SELECT t.AccountId, t.AccountCode, t.AccountName, t.ParentAccountCode, t.isHead FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead, (select isnull(sum(LocalCurrencyAmount), 0) from GLTransactions where GLTransactions.AccountCode = ret.AccountCode and GLTransactions.AddedDate >= '" + pGainLossModel.FromDate.ToString("MM/dd/yyyy") + "' and GLTransactions.AddedDate <= '" + pGainLossModel.ToDate.ToString("MM/dd/yyyy") + "') as bal FROM ret where isHead = 1";
        //    DataTable dtOtherExpenses = DBUtils.GetDataTable(SqlOtherExpenses);
        //    if (dtOtherExpenses != null && dtOtherExpenses.Rows.Count > 0)
        //    {
        //        table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Other Expenses</b></h5></td><td></td><td></td></tr>";
        //        foreach (DataRow dr4 in dtOtherExpenses.Rows)
        //        {
        //            if (Common.toDecimal(dr4["bal"]) == 0)
        //            {
        //                if (pGainLossModel.ShowAll)
        //                {
        //                    table += "<tr><td>" + dr4["AccountName"].ToString() + "</td>";
        //                    Balance5 = Math.Round(Common.toDecimal(dr4["bal"]), 2);
        //                    TotalOtherExpenses += Balance5;
        //                    table += "<td></td><td align='right'>" + Balance5 + "</td></tr>";
        //                }
        //            }
        //            else
        //            {
        //                table += "<tr><td>" + dr4["AccountName"].ToString() + "</td>";
        //                Balance5 = Math.Round(Common.toDecimal(dr4["bal"]), 2);
        //                TotalOtherExpenses += Balance5;
        //                table += "<td></td><td align='right'>" + Balance5 + "</td></tr>";
        //            }
        //        }
        //        table += "<tr><td><b>Total other expenses</b></td>";
        //        table += "<td></td><td align='right'><b>" + TotalOtherExpenses + "</b></td></tr>";
        //        table += "<tr><td height='30'></td></tr>";
        //    }

        //    NetProfit = ((GrossProfit + 0) + (TotalOtherIncome + 0) + (TotalOtherExpenses + 0)) - (TotalAdministrativeCost + 0);
        //    table += "<tr style='border-bottom:#000 1px solid; border-top:#000 1px solid;'><td><b>Net profit</b></td><td></td>";
        //    table += "<td align='right'><b>" + NetProfit + "</b></td></tr>";
        //    table += "</tbody>";
        //    table += "</table>";

        //    pGainLossModel.HtmlTable = table;
        //    return pGainLossModel;
        //}

        //public static XLWorkbook BuildIncomeTable1(ChartOfAccountsModel pGainLossModel)
        //{
        //    decimal Balance1 = 0.0m;
        //    decimal Balance2 = 0.0m;
        //    decimal Balance3 = 0.0m;
        //    decimal Balance4 = 0.0m;
        //    decimal Balance5 = 0.0m;
        //    decimal TotalIncome = 0.0m;
        //    decimal TotalCostofSales = 0.0m;
        //    decimal TotalOtherIncome = 0.0m;
        //    decimal TotalAdministrativeCost = 0.0m;
        //    decimal TotalOtherExpenses = 0.0m;
        //    decimal GrossProfit = 0.0m;
        //    decimal NetProfit = 0.0m;
        //    string RetVal = "";
        //    int mycount1 = 1;
        //    pGainLossModel.TotalBalance = 0;
        //    pGainLossModel.tempTable = "";
        //    XLWorkbook workbook = new XLWorkbook();
        //    var worksheet = workbook.Worksheets.Add("Profit and Loss Report");
        //    //worksheet.Cell("A1").Value = "Profit and Loss Report";

        //    //var worksheet = workbook.Worksheet(1);

        //    string SqlIncome = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossIncomeAccount + "'  UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
        //    DataTable dtIncome = DBUtils.GetDataTable(SqlIncome);
        //    if (dtIncome != null && dtIncome.Rows.Count > 0)
        //    {
        //        TotalIncome = GetBalance(dtIncome.Rows[0]["AccountName"].ToString(), pGainLossModel.FromDate, pGainLossModel.ToDate);
        //        foreach (DataRow dr in dtIncome.Rows)
        //        {
        //            Balance1 = GetBalance(dr["AccountName"].ToString(), pGainLossModel.FromDate, pGainLossModel.ToDate);

        //            if (Balance1 == 0)
        //            {
        //                if (pGainLossModel.ShowAll)
        //                {
        //                    mycount1++;
        //                    worksheet.Cell("A" + mycount1).Value = dr["AccountName"].ToString();
        //                    TotalIncome += Balance1;
        //                    worksheet.Cell("C" + mycount1).Value = Balance1;
        //                }
        //            }
        //            else
        //            {
        //                if (Balance1 > 0)
        //                {
        //                    mycount1++;
        //                    worksheet.Cell("A" + mycount1).Value = dr["AccountName"].ToString();
        //                    Balance1 = Balance1 * (-1);
        //                    TotalIncome += Balance1;
        //                    worksheet.Cell("C" + mycount1).Value = Balance1;
        //                }
        //                else
        //                {
        //                    mycount1++;
        //                    worksheet.Cell("A" + mycount1).Value = dr["AccountName"].ToString();
        //                    Balance1 = Math.Abs(Balance1);
        //                    TotalIncome += Balance1;
        //                    worksheet.Cell("C" + mycount1).Value = Balance1;
        //                }
        //            }
        //        }
        //    }
        //    return workbook;
        //}
        public void ExportToExcelGainAndLoss(ChartOfAccountsModel pGainLossModel)
        {
            try
            {
                decimal Balance1 = 0.0m;
                decimal Balance2 = 0.0m;
                decimal Balance3 = 0.0m;
                decimal Balance4 = 0.0m;
                decimal Balance5 = 0.0m;
                decimal TotalIncome = 0.0m;
                decimal TotalCostofSales = 0.0m;
                decimal TotalOtherIncome = 0.0m;
                decimal TotalAdministrativeCost = 0.0m;
                decimal TotalOtherExpenses = 0.0m;
                decimal GrossProfit = 0.0m;
                decimal NetProfit = 0.0m;
                int level = 0;
                string spaces = "";
                XLWorkbook workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Profit and Loss Report");
                worksheet.Cell("A1").Value = "Profit and Loss Report";

                worksheet.Cell("A1").Style.Font.Bold = true;
                worksheet.Cell("A1").Style.Font.FontSize = 20;
                worksheet.Cell("A2").Value = "Profit and Loss";
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("B2").Value = "From : " + pGainLossModel.FromDate.ToShortDateString() + "";
                worksheet.Cell("C2").Value = "To : " + pGainLossModel.ToDate.ToShortDateString() + "";
                worksheet.Cell("E2").Style.Font.Bold = true;

                worksheet.Cell("A3").Value = "";
                worksheet.Cell("C3").Value = "                                              " + SysPrefs.DefaultCurrency;
                worksheet.Cell("C3").Style.Font.Bold = true;
                worksheet.Cell("A3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("B3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("C3").Style.Fill.BackgroundColor = XLColor.PersianRed;

                worksheet.Cell("A3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("B3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("C3").Style.Font.FontColor = XLColor.White;

                worksheet.ColumnWidth = 25;
                int mycount1 = 4;
                worksheet.Cell("A" + mycount1).Value = "Income";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;
                string SqlIncome = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossIncomeAccount + "'  UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";

                DataTable dtIncome = DBUtils.GetDataTable(SqlIncome);
                if (dtIncome != null && dtIncome.Rows.Count > 0)
                {
                    TotalIncome = GetBalance(SysPrefs.ProfileNLossIncomeAccount, pGainLossModel.FromDate, pGainLossModel.ToDate);
                    TotalIncome = Math.Abs(TotalIncome);// for income only
                    foreach (DataRow dr in dtIncome.Rows)
                    {
                        bool isHead = Common.toBool(dr["isHead"]);
                        spaces = "";
                        level = Common.toInt(dr["level"]);
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }

                        Balance1 = GetBalance(dr["AccountCode"].ToString(), pGainLossModel.FromDate, pGainLossModel.ToDate);

                        //TotalIncome += Balance1;
                        if (Balance1 > 0)
                        {
                            Balance1 = -1 * Balance1;
                        }
                        else
                        {
                            Balance1 = Math.Abs(Balance1);
                        }
                        TotalIncome += Balance1;
                        if (Balance1 == 0)
                        {
                            if (isHead)
                            {
                                if (pGainLossModel.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + dr["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Balance1;
                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Balance1;
                            }

                        }
                        else
                        {
                            if (Balance1 > 0)
                            {
                                if (isHead)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + dr["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Balance1;
                                }
                                else
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + dr["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Balance1;
                                }
                                //mycount1++;
                                //worksheet.Cell("A" + mycount1).Value = spaces + dr["AccountName"].ToString();
                                //Balance1 = Balance1 * (-1);
                                //worksheet.Cell("C" + mycount1).Value = Balance1;
                            }
                            else
                            {
                                //mycount1++;
                                //worksheet.Cell("A" + mycount1).Value = spaces + dr["AccountName"].ToString();
                                //Balance1 = Math.Abs(Balance1);
                                //worksheet.Cell("C" + mycount1).Value = Balance1;
                                if (isHead)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + dr["AccountName"].ToString();
                                    //Balance1 = Math.Abs(Balance1);
                                    //worksheet.Cell("C" + mycount1).Value = Math.Abs(Balance1);
                                    worksheet.Cell("C" + mycount1).Value = Balance1;
                                }
                                else
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + dr["AccountName"].ToString();
                                    // Balance1 = Math.Abs(Balance1);
                                    //worksheet.Cell("C" + mycount1).Value = Math.Abs(Balance1);
                                    worksheet.Cell("C" + mycount1).Value = Balance1;
                                }
                            }
                        }
                    }
                }
                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total income";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalIncome;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 2;

                worksheet.Cell("A" + mycount1).Value = "Cost of sales";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;
                //table += "<tr style='border-bottom:#000 1px solid;'><td></td><td></td><td></td></tr>";
                //string SqlCostofSales = "; WITH ret AS(SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossCostOfSales + "' UNION ALL SELECT t.AccountId, t.AccountCode, t.AccountName, t.ParentAccountCode, t.isHead FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead, (select isnull(sum(LocalCurrencyAmount), 0) from GLTransactions where GLTransactions.AccountCode = ret.AccountCode and GLTransactions.AddedDate >= '" + pGainLossModel.FromDate.ToString("MM/dd/yyyy") + "' and GLTransactions.AddedDate <= '" + pGainLossModel.ToDate.ToString("MM/dd/yyyy") + "') as bal FROM ret where isHead = 1";
                string SqlCostofSales = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossCostOfSales + "'  UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
                DataTable dtCostofSales = DBUtils.GetDataTable(SqlCostofSales);
                if (dtCostofSales != null && dtCostofSales.Rows.Count > 0)
                {
                    TotalCostofSales = GetBalance(SysPrefs.ProfileNLossCostOfSales, pGainLossModel.FromDate, pGainLossModel.ToDate);
                    foreach (DataRow dr1 in dtCostofSales.Rows)
                    {
                        bool isHead = Common.toBool(dr1["isHead"]);
                        spaces = "";
                        level = Common.toInt(dr1["level"]);
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }

                        Balance2 = GetBalance(dr1["AccountCode"].ToString(), pGainLossModel.FromDate, pGainLossModel.ToDate);
                        TotalCostofSales += Balance2;
                        TotalCostofSales += Math.Abs(TotalCostofSales);

                        if (Balance2 == 0)
                        {
                            if (isHead)
                            {
                                if (pGainLossModel.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + dr1["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Balance2;
                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr1["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Balance2;
                            }
                        }
                        else
                        {
                            if (Balance2 > 0)
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr1["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Balance2;
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr1["AccountName"].ToString();
                                Balance1 = Math.Abs(Balance2);
                                worksheet.Cell("C" + mycount1).Value = "(" + Balance2 + ")";
                            }
                        }
                    }
                }
                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total Cost of sales";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalCostofSales;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 2;

                GrossProfit = TotalIncome - TotalCostofSales;
                worksheet.Cell("A" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("A" + mycount1).Value = "Gross Profit";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("B" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("C" + mycount1).Value = GrossProfit;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 3;
                worksheet.Cell("A" + mycount1).Value = "Other income";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;

                //SqlOtherIncomestring SqlOtherIncome = "; WITH ret AS(SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossOtherIncome + "' UNION ALL SELECT t.AccountId, t.AccountCode, t.AccountName, t.ParentAccountCode, t.isHead FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead, (select isnull(sum(LocalCurrencyAmount), 0) from GLTransactions where GLTransactions.AccountCode = ret.AccountCode and GLTransactions.AddedDate >= '" + pGainLossModel.FromDate.ToString("MM/dd/yyyy") + "' and GLTransactions.AddedDate <= '" + pGainLossModel.ToDate.ToString("MM/dd/yyyy") + "') as bal FROM ret where isHead = 1";
                string SqlOtherIncome = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossOtherIncome + "'  UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
                DataTable dtOtherIncome = DBUtils.GetDataTable(SqlOtherIncome);
                if (dtOtherIncome != null && dtOtherIncome.Rows.Count > 0)
                {
                    TotalOtherIncome = GetBalance(SysPrefs.ProfileNLossOtherIncome, pGainLossModel.FromDate, pGainLossModel.ToDate);
                    foreach (DataRow dr2 in dtOtherIncome.Rows)
                    {
                        bool isHead = Common.toBool(dr2["isHead"]);
                        spaces = "";
                        level = Common.toInt(dr2["level"]);
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }

                        Balance3 = GetBalance(dr2["AccountCode"].ToString(), pGainLossModel.FromDate, pGainLossModel.ToDate);
                        TotalOtherIncome += Balance3;
                        TotalOtherIncome = Math.Abs(TotalOtherIncome);


                        if (Balance3 == 0)
                        {
                            if (isHead)
                            {
                                if (pGainLossModel.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + dr2["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Balance3;
                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr2["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Balance3;
                            }

                        }
                        else
                        {
                            if (Balance3 > 0)
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr2["AccountName"].ToString();
                                // Balance3 = Balance3 * (-1);
                                worksheet.Cell("C" + mycount1).Value = Balance3;
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr2["AccountName"].ToString();
                                Balance3 = Math.Abs(Balance3);
                                worksheet.Cell("C" + mycount1).Value = "(" + Balance3 + ")";
                            }
                        }
                    }
                }
                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total Other income";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalOtherIncome;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 2;
                worksheet.Cell("A" + mycount1).Value = "Administrative costs";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;

                //string SqlAdministrativeCost = "; WITH ret AS(SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossAdministrativeExpenses + "' UNION ALL SELECT t.AccountId, t.AccountCode, t.AccountName, t.ParentAccountCode, t.isHead FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead, (select isnull(sum(LocalCurrencyAmount), 0) from GLTransactions where GLTransactions.AccountCode = ret.AccountCode and GLTransactions.AddedDate >= '" + pGainLossModel.FromDate.ToString("MM/dd/yyyy") + "' and GLTransactions.AddedDate <= '" + pGainLossModel.ToDate.ToString("MM/dd/yyyy") + "') as bal FROM ret where isHead = 1";
                string SqlAdministrativeCost = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossAdministrativeExpenses + "'  UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
                DataTable dtAdministrativeCost = DBUtils.GetDataTable(SqlAdministrativeCost);
                if (dtAdministrativeCost != null && dtAdministrativeCost.Rows.Count > 0)
                {
                    TotalAdministrativeCost = GetBalance(SysPrefs.ProfileNLossAdministrativeExpenses, pGainLossModel.FromDate, pGainLossModel.ToDate);
                    foreach (DataRow dr3 in dtAdministrativeCost.Rows)
                    {
                        bool isHead = Common.toBool(dr3["isHead"]);
                        spaces = "";
                        level = Common.toInt(dr3["level"]);
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }

                        Balance4 = GetBalance(dr3["AccountCode"].ToString(), pGainLossModel.FromDate, pGainLossModel.ToDate);
                        TotalAdministrativeCost += Balance4;
                        TotalAdministrativeCost = Math.Abs(TotalAdministrativeCost);


                        if (Balance4 == 0)
                        {
                            if (isHead)
                            {
                                if (pGainLossModel.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + dr3["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Balance4;
                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr3["AccountName"].ToString();
                                //  worksheet.Cell("C" + mycount1).Value = Balance5;
                                worksheet.Cell("C" + mycount1).Value = Balance4;
                            }
                        }
                        else
                        {
                            if (Balance4 > 0)
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr3["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Balance4;
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr3["AccountName"].ToString();
                                Balance4 = Math.Abs(Balance4);
                                worksheet.Cell("C" + mycount1).Value = "(" + Balance4 + ")";
                            }
                        }

                    }
                }
                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total Administrative costs";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalAdministrativeCost;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;



                mycount1 = mycount1 + 2;
                worksheet.Cell("A" + mycount1).Value = "Other Expenses";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;

                //string SqlOtherExpenses = "; WITH ret AS(SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossOtherExpenses + "' UNION ALL SELECT t.AccountId, t.AccountCode, t.AccountName, t.ParentAccountCode, t.isHead FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountId, AccountCode, AccountName, ParentAccountCode, isHead, (select isnull(sum(LocalCurrencyAmount), 0) from GLTransactions where GLTransactions.AccountCode = ret.AccountCode and GLTransactions.AddedDate >= '" + pGainLossModel.FromDate.ToString("MM/dd/yyyy") + "' and GLTransactions.AddedDate <= '" + pGainLossModel.ToDate.ToString("MM/dd/yyyy") + "') as bal FROM ret where isHead = 1";
                string SqlOtherExpenses = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.ProfileNLossOtherExpenses + "'  UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
                DataTable dtOtherExpenses = DBUtils.GetDataTable(SqlOtherExpenses);
                if (dtOtherExpenses != null && dtOtherExpenses.Rows.Count > 0)
                {
                    TotalOtherExpenses = GetBalance(SysPrefs.ProfileNLossOtherExpenses, pGainLossModel.FromDate, pGainLossModel.ToDate);

                    foreach (DataRow dr4 in dtOtherExpenses.Rows)
                    {
                        bool isHead = Common.toBool(dr4["isHead"]);
                        spaces = "";
                        level = Common.toInt(dr4["level"]);
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }
                        Balance5 = GetBalance(dr4["AccountCode"].ToString(), pGainLossModel.FromDate, pGainLossModel.ToDate);
                        TotalOtherExpenses += Balance5;
                        TotalOtherExpenses = Math.Round(TotalOtherExpenses);
                        if (Balance5 == 0)
                        {
                            if (isHead)
                            {
                                if (pGainLossModel.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + dr4["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Balance5;
                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr4["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Balance5;
                            }

                        }
                        else
                        {
                            if (Balance5 > 0)
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr4["AccountName"].ToString();
                                Balance5 = GetBalance(dr4["AccountCode"].ToString(), pGainLossModel.FromDate, pGainLossModel.ToDate);
                                // TotalOtherExpenses += Balance5;
                                worksheet.Cell("C" + mycount1).Value = Balance5;
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + dr4["AccountName"].ToString();
                                Balance5 = GetBalance(dr4["AccountCode"].ToString(), pGainLossModel.FromDate, pGainLossModel.ToDate);
                                Balance5 = Math.Abs(Balance5);
                                // TotalOtherExpenses += Balance5;
                                worksheet.Cell("C" + mycount1).Value = "(" + Balance5 + ")";
                            }
                        }
                    }
                    //Total other expenses
                }
                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total other expenses";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalOtherExpenses;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                decimal NetProfitA = GrossProfit + Math.Abs(TotalOtherIncome);
                decimal NetProfitB = NetProfitA - TotalAdministrativeCost;
                NetProfit = NetProfitB - Math.Abs(TotalOtherExpenses);
                // NetProfit = ((GrossProfit + 0) + (TotalOtherIncome + 0) + (TotalOtherExpenses + 0)) - (TotalAdministrativeCost + 0);
                mycount1 = mycount1 + 2;
                worksheet.Cell("A" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("A" + mycount1).Value = "Net profit";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("C" + mycount1).Value = NetProfit;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;

                worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.White;

                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("B" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                string fileName = Common.getRandomDigit(4) + "-" + "AccountBalanceGainAndLossReport";
                workbook.SaveAs(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".xlsx"));
                Response.Clear();
                Response.ClearHeaders();
                Response.ClearContent();
                Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".xlsx");
                Response.ContentType = "text/plain";
                Response.Flush();
                //Response.TransmitFile(Server.MapPath("~/Uploads/" + "/1/DeanRpt.xlsx"));
                Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".xlsx"));
                Response.End();

            }
            catch (Exception x)
            {
                Response.Write("<font color='Red'>" + x.Message + " </font>");
            }

        }

        //public void ExportToPDFGainAndLoss(ChartOfAccountsModel pGainLossModel)
        //{
        //    string HostName = System.Web.HttpContext.Current.Request.Url.Host;
        //    pGainLossModel.isPrint = true;
        //    pGainLossModel = BalanceGainAndLossTable(pGainLossModel);
        //    string fileName = Common.getRandomDigit(4) + "-" + "AccountBalanceGainAndLossReport";
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/");
        //    HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //    htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //    PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(pGainLossModel.HtmlTable, sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    Response.Clear();
        //    Response.ClearContent();
        //    Response.ClearHeaders();
        //    Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".pdf");
        //    Response.ContentType = "application/pdf";
        //    Response.Flush();
        //    Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".pdf"));
        //    Response.End();
        //}
        #endregion

        #region Balance Sheet

        public ActionResult BalanceSheetReport()
        {
            ChartOfAccountsModel pBalanceSheet = new ChartOfAccountsModel();
            pBalanceSheet.FromDate = Common.toDateTime(SysPrefs.PostingDate).AddDays(-30); ;
            pBalanceSheet.ToDate = Common.toDateTime(SysPrefs.PostingDate);
            pBalanceSheet.isvalid = false;
            pBalanceSheet.ShowAll = false;
            return View("~/Views/Reports/Accounts/BalanceSheetReport.cshtml", pBalanceSheet);
        }
        [HttpPost]
        public ActionResult BalanceSheetReport(ChartOfAccountsModel pBalanceSheet)
        {
            pBalanceSheet = BalanceSheetReportTable(pBalanceSheet);
            return View("~/Views/Reports/Accounts/BalanceSheetReport.cshtml", pBalanceSheet);
        }
        public static decimal GetBalance(string pAccountCode, DateTime FromDate, DateTime ToDate)
        {
            decimal Balance = 0;
            if (pAccountCode.Trim() != "")
            {
                string strBalanceQury = "  select isNull( SUM(GLTransactions.LocalCurrencyAmount), 0) as LCBalance   From GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID Where GLReferences.isPosted=1 And  AccountCode = '" + pAccountCode + "' And AddedDate >='" + Common.toDateTime(FromDate).ToString("yyyy-MM-dd 00:00:00") + "' and AddedDate<= '" + Common.toDateTime(ToDate).ToString("yyyy-MM-dd 23:59:00") + "' ";
                //strBalanceQury = ";WITH ret AS(SELECT  AccountCode FROM GLChartOfAccounts WHERE AccountCode = '" + pAccountCode + "' UNION ALL SELECT  t.AccountCode FROM    GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT  isNull( SUM(GLTransactions.LocalCurrencyAmount), 0) as LCBalance From GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID Where GLReferences.isPosted=1 And GLTransactions.AccountCode in (SELECT  AccountCode FROM ret) And GLTransactions.AddedDate >='" + Common.toDateTime(FromDate).ToString("yyyy-MM-dd 00:00:00") + "' and AddedDate<= '" + Common.toDateTime(ToDate).ToString("yyyy-MM-dd 23:59:00") + "' ";
                string strBalance = DBUtils.executeSqlGetSingle(strBalanceQury);
                Balance = Math.Round(Common.toDecimal(strBalance), 2);
            }
            return Balance;
        }
        public static ChartOfAccountsModel BuildBSTable(ChartOfAccountsModel pBalanceSheet)
        {
            string RetVal = "";
            string spaces = "";
            pBalanceSheet.TotalBalance = 0;
            pBalanceSheet.tempTable = "";
            string sqlQuery = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + pBalanceSheet.AccountCode + "' UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
            DataTable dtAccount = DBUtils.GetDataTable(sqlQuery);
            if (dtAccount != null)
            {
                if (dtAccount.Rows.Count > 0)
                {
                    //pBalanceSheet.TotalBalance = GetBalance(Common.toString(dtAccount.Rows[0]["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                    pBalanceSheet.TotalBalance = GetBalance(pBalanceSheet.AccountCode, pBalanceSheet.FromDate, pBalanceSheet.ToDate);

                }
                foreach (DataRow drAccount in dtAccount.Rows)
                {
                    if (Common.toString(pBalanceSheet.AccountCode).Trim() != Common.toString(drAccount["AccountCode"]).Trim())
                    {
                        decimal level = 0;
                        decimal Currbalance = GetBalance(Common.toString(drAccount["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                        level = Common.toDecimal(drAccount["level"]);
                        bool isHead = Common.toBool(drAccount["isHead"]);
                        //RetVal += "<br>" + Currbalance + "-" + level + "-" + isHead + "-" + Common.toString(drAccount["AccountCode"]) + "-" + pBalanceSheet.FromDate + "-" + pBalanceSheet.ToDate;
                        pBalanceSheet.TotalBalance += Currbalance;
                        //pBalanceSheet.TotalBalance = Math.Abs(pBalanceSheet.TotalBalance);
                        if (level == 0)
                        {
                            if (Common.toDecimal(Currbalance) == 0)
                            {
                                if (isHead)
                                {
                                    if (pBalanceSheet.ShowAll)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                }
                                else
                                {
                                    RetVal += "<tr style='background-color:#dff0d8;'  data-depth=" + level + " class='collapsen level" + level + "'>";
                                    RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                    RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                }
                            }
                            else
                            {
                                if (Common.toDecimal(Currbalance) > 0)
                                {
                                    if (isHead)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                    else
                                    {
                                        RetVal += "<tr style='background-color:#dff0d8;'  data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                }
                                else
                                {
                                    if (isHead)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + " </td></tr>";
                                    }
                                    else
                                    {
                                        RetVal += "<tr style='background-color:#dff0d8;' data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + "</td></tr>";
                                    }
                                }
                            }
                        }
                        else if (level == 1)
                        {
                            if (Common.toDecimal(Currbalance) == 0)
                            {
                                if (isHead)
                                {
                                    if (pBalanceSheet.ShowAll)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                }
                                else
                                {

                                    RetVal += "<tr style='background-color: #E8FfFF;'  data-depth=" + level + " class='collapsen level" + level + "'>";
                                    RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                    RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                }
                            }
                            else
                            {
                                if (Common.toDecimal(Currbalance) > 0)
                                {
                                    if (isHead)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                    else
                                    {
                                        RetVal += "<tr style='background-color: #E8FfFF;'  data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                }
                                else
                                {
                                    if (isHead)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + " </td></tr>";
                                    }
                                    else
                                    {
                                        RetVal += "<tr style='background-color: #E8FfFF;' data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + "</td></tr>";
                                    }
                                }
                            }
                        }
                        else if (level == 2)
                        {
                            if (Common.toDecimal(Currbalance) == 0)
                            {
                                if (isHead)
                                {
                                    if (pBalanceSheet.ShowAll)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                }
                                else
                                {

                                    RetVal += "<tr style='background-color: #edf2f5;'  data-depth=" + level + " class='collapsen level" + level + "'>";
                                    RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                    RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                }
                            }
                            else
                            {
                                if (Common.toDecimal(Currbalance) > 0)
                                {
                                    if (isHead)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                    else
                                    {
                                        RetVal += "<tr style='background-color:  #edf2f5;'  data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                }
                                else
                                {
                                    if (isHead)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + " </td></tr>";
                                    }
                                    else
                                    {
                                        RetVal += "<tr style='background-color:  #edf2f5;' data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + "</td></tr>";
                                    }
                                }
                            }
                        }
                        else if (level == 3)
                        {
                            if (Common.toDecimal(Currbalance) == 0)
                            {
                                if (isHead)
                                {
                                    if (pBalanceSheet.ShowAll)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                }
                                else
                                {

                                    RetVal += "<tr style='background-color: #fcf8e3;'  data-depth=" + level + " class='collapsen level" + level + "'>";
                                    RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                    RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                }
                            }
                            else
                            {
                                if (Common.toDecimal(Currbalance) > 0)
                                {
                                    if (isHead)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                    else
                                    {
                                        RetVal += "<tr style='background-color: #fcf8e3;'  data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + Currbalance + "</td></tr>";
                                    }
                                }
                                else
                                {
                                    if (isHead)
                                    {
                                        RetVal += "<tr data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td>" + "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + " </td></tr>";
                                    }
                                    else
                                    {
                                        RetVal += "<tr style='background-color: #fcf8e3;' data-depth=" + level + " class='collapsen level" + level + "'>";
                                        RetVal += "<td><span class='toggle collapsen'></span>" + "&nbsp;&nbsp;" + drAccount["AccountName"] + "</td>";
                                        RetVal += "<td></td><td align='right'>" + "(" + Math.Abs(Currbalance) + ")" + "</td></tr>";
                                    }
                                }
                            }
                        }
                        pBalanceSheet.tempTable = RetVal;
                    }
                }
            }

            return pBalanceSheet;
        }

        public static ChartOfAccountsModel BalanceSheetReportTable(ChartOfAccountsModel pBalanceSheet)
        {
            pBalanceSheet.isvalid = true;
            decimal TotalTangibleAssets = 0.0m;
            decimal TotalCurrentAssets = 0.0m;
            decimal TotalCurrentLiabilities = 0.0m;
            decimal TotalNonCurrentLiabilities = 0.0m;
            decimal TotalCapitalandReserves = 0.0m;
            decimal TotalAssets = 0.0m;
            decimal TotalCapitalandLiabilites = 0.0m;
            decimal TotalAssetsLessCurrentLiabilities = 0.0m;

            #region Table Header
            string table = "";
            table += "<div class='row'>";
            table += "<div class='col-md-6'>";
            table += "<div class='media'>";
            table += "<div class='media -body'>";
            table += "<h4 align='center' class='text-center'><b>Balance Sheet</b></h4>";
            table += "</div>";
            table += "</div>";
            table += "<br/><br/>";
            table += "<div align='center' class='text-center' style='color:darkred;'><strong>" + @SysPrefs.SiteName + "</strong></div>";
            table += "<br/><br/>";
            table += "<h4><b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; As of:&nbsp; &nbsp;" + pBalanceSheet.FromDate.ToShortDateString() + "&nbsp; &nbsp; &nbsp; &nbsp;</b></h4>";
            table += "</div>";
            table += "</div>";
            table += "<table border='0' width='95%' cellpadding='5'>";
            table += "<thead>";
            table += " <tr>";
            table += " </tr>";
            table += " <tr>";
            table += "<th style='background-color:#1c3a70; color:#FFF;'></th>";
            table += "<th style='background-color:#1c3a70; color:#FFF; text-align:right;'></th>";
            table += "<th style='background-color:#1c3a70; color:#FFF; text-align:right;'>GBP</th>";
            table += "</tr>";
            table += "<thead>";
            table += "<tbody>";
            table += "<tr style='border-bottom:#000 1px solid;'>";
            table += "<td><h5><b>Non-Current Assets</b><h5></td><td></td><td></td></tr>";
            //table += "<tr><td><b> Tangible Assets </b></td>";
            //table += "<td></td><td></td></tr>";

            #endregion
            pBalanceSheet.FromDate = pBalanceSheet.FromDate;
            pBalanceSheet.ToDate = PostingUtils.BeginFiscalYear();

            DateTime FromDate = Common.toDateTime(pBalanceSheet.FromDate.ToString("MM/dd/yyyy"));
            DateTime ToDate = Common.toDateTime(pBalanceSheet.ToDate.ToString("MM/dd/yyyy"));

            #region Non Current Assets 
            table += "<table border='0' id='tblNonCurrentAssests' width='95%' cellpadding='5'>";
            pBalanceSheet.AccountCode = SysPrefs.BSNonCurrentAssetsAccount;
            pBalanceSheet = BuildBSTable(pBalanceSheet);
            table += pBalanceSheet.tempTable;
            TotalTangibleAssets = pBalanceSheet.TotalBalance;
            table += "</table>";
            table += "<table border='0' width='95%' cellpadding='5'>";
            //table += BuildBSTable(SysPrefs.BSNonCurrentAssetsAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            table += "<tr><td><b>Total Tangible Assets</b></td>";
            if (TotalTangibleAssets > 0)
            {
                table += "<td></td><td align='right'><b>" + TotalTangibleAssets + "</b></td></tr>";
            }
            else
            {
                table += "<td></td><td align='right'><b>(" + Math.Abs(TotalTangibleAssets) + ")</b></td></tr>";
            }


            //table += "<td></td><td align='right'><b>" + TotalTangibleAssets + "</b></td></tr>";

            table += "<tr><td height='30'></td></tr>";
            table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Current Assets</b></h5></td><td></td><td></td></tr>";
            //table += "<tr><td><b> Cash at bank and in hand </b></td>";
            //table += "<td></td><td></td></tr>";
            table += "</table>";
            #endregion

            #region Current Assests
            table += "<table border='0' id='tblCurrentAssests' width='95%' cellpadding='5'>";
            pBalanceSheet.AccountCode = SysPrefs.BSCurrentAssetsAccount;
            pBalanceSheet = BuildBSTable(pBalanceSheet);
            table += pBalanceSheet.tempTable;
            TotalCurrentAssets = pBalanceSheet.TotalBalance;
            table += "</table>";
            table += "<table border='0' width='95%' cellpadding='5'>";
            //table += BuildBSTable(SysPrefs.BSCurrentAssetsAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            table += "<tr><td><b>Total Current Assets</b></td>";
            table += "<td></td><td align='right'><b>" + TotalCurrentAssets + "</b></td></tr>";

            table += "<tr><td height='30'></td></tr>";

            TotalAssets = TotalTangibleAssets + TotalCurrentAssets;
            table += "<tr style='border-bottom:#000 1px solid; border-top:#000 1px solid;'><td><b>Total Assets</b></td><td></td>";
            table += "<td align='right'><b>" + TotalAssets + "</b></td></tr>";
            table += "<tr><td height='60'></td></tr>";
            table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Liabilities</b></h5></td><td></td><td></td></tr>";
            table += "</table>";
            #endregion

            #region Current Liabilities
            table += "<table border='0' id='tblCurrentLiabilities' width='95%' cellpadding='5'>";
            pBalanceSheet.AccountCode = SysPrefs.BSCurrentLiabilitesAccount;
            pBalanceSheet = BuildBSTable(pBalanceSheet);
            table += pBalanceSheet.tempTable;
            TotalCurrentLiabilities = pBalanceSheet.TotalBalance;
            table += "</table>";
            table += "<table border='0' width='95%' cellpadding='5'>";
            //table += BuildBSTable(SysPrefs.BSCurrentLiabilitesAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            //table += "<tr><td><b>Liabilities</b></td>";
            //table += "<td></td><td align='right'><b>" + TotalCurrentLiabilities + "</b></td></tr>";
            TotalAssetsLessCurrentLiabilities = TotalTangibleAssets + (TotalAssets - TotalCurrentLiabilities);
            table += "<tr><td height='30'></td></tr>";
            table += "<tr style='border-bottom:#000 1px solid; border-top:#000 1px solid;'><td><b>Total Current Liabilities</b></td><td></td>";
            table += "<td align='right'><b>" + TotalCurrentLiabilities + "</b></td></tr>";

            table += "<tr><td height='30'></td></tr>";
            table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Capital and Reserves</b></h5></td><td></td><td></td></tr>";
            table += "</table>";
            #endregion

            //#region Non Current Liabilities
            //table += "<table border='0' id='tblNonCurrentLiabilities' width='95%' cellpadding='5'>";
            //pBalanceSheet.AccountCode = SysPrefs.BSNonCurrentliabilitesAccount;
            //pBalanceSheet = BuildBSTable(pBalanceSheet);
            //table += pBalanceSheet.tempTable;
            //TotalNonCurrentLiabilities = pBalanceSheet.TotalBalance;
            //table += "</table>";
            //table += "<table border='0' width='95%' cellpadding='5'>";
            ////table += BuildBSTable(SysPrefs.BSNonCurrentliabilitesAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            //table += "<tr><td><b>Total Creditors: amounts falling due after more than one year</b></td>";
            //table += "<td></td><td align='right'><b>" + TotalNonCurrentLiabilities + "</b></td></tr>";
            //table += "<tr><td height='60'></td></tr>";
            //table += "<tr style='border-bottom:#000 1px solid;'><td><h5><b>Capital and Reserves</b></h5></td><td></td><td></td></tr>";
            //table += "</table>";
            //#endregion

            #region  Capital and Reserves
            table += "<table border='0' id='tblCapitalandReserves' width='95%' cellpadding='5'>";
            pBalanceSheet.AccountCode = SysPrefs.BSCapitalNReservesAccount;
            pBalanceSheet = BuildBSTable(pBalanceSheet);
            table += pBalanceSheet.tempTable;
            TotalCapitalandReserves = pBalanceSheet.TotalBalance;
            table += "</table>";
            table += "<table border='0' width='95%' cellpadding='5'>";
            //table += BuildBSTable(SysPrefs.BSCapitalNReservesAccount, FromDate, ToDate, pBalanceSheet.ShowAll);
            table += "<tr><td><b>Total Capital and Reserves</b></td>";
            table += "<td></td><td align='right'><b>" + TotalCapitalandReserves + "</b></td></tr>";

            table += "<tr><td height='30'></td></tr>";
            TotalCapitalandLiabilites = TotalNonCurrentLiabilities + TotalCapitalandReserves;
            table += "<tr style='border-bottom:#000 1px solid; border-top:#000 1px solid;'><td><b>Total Equity and Liabilites</b></td><td></td>";
            //table += "<td align='right'><b>" + TotalCapitalandLiabilites + "</b></td></tr>";
            table += "<td align='right'><b>" + (TotalCurrentLiabilities + TotalCapitalandLiabilites - TotalAssets).CurrencySeparator() + "</b></td></tr>";
            table += "</table>";
            #endregion

            table += "</tbody>";
            table += "</table>";
            pBalanceSheet.HtmlTable = table;
            return pBalanceSheet;
        }


        //public void BalanceSheetExportToPDF(ChartOfAccountsModel pBalanceSheet)
        //{
        //    string HostName = System.Web.HttpContext.Current.Request.Url.Host;
        //    pBalanceSheet = BalanceSheetReportTable(pBalanceSheet);
        //    string fileName = Common.getRandomDigit(4) + "-" + "BalanceSheet";
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/");
        //    HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //    htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //    PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(pBalanceSheet.HtmlTable, sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    Response.Clear();
        //    Response.ClearContent();
        //    Response.ClearHeaders();
        //    Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".pdf");
        //    Response.ContentType = "application/pdf";
        //    Response.Flush();
        //    Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".pdf"));
        //    Response.End();
        //}

        public void BalanceSheetExportToExcel(ChartOfAccountsModel pBalanceSheet)
        {
            try
            {
                pBalanceSheet.isvalid = true;
                decimal TotalTangibleAssets = 0.0m;
                decimal TotalCurrentAssets = 0.0m;
                decimal TotalCurrentLiabilities = 0.0m;
                decimal TotalNonCurrentLiabilities = 0.0m;
                decimal TotalAssetsLessCurrentLiabilities = 0.0m;
                decimal TotalCapitalandReserves = 0.0m;
                decimal TotalAssets = 0.0m;
                decimal TotalCapitalandLiabilites = 0.0m;
                string spaces = "";
                int level = 0;
                pBalanceSheet.ToDate = pBalanceSheet.FromDate;
                pBalanceSheet.FromDate = PostingUtils.BeginFiscalYear();

                XLWorkbook workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Balance Sheet");
                worksheet.Cell("A1").Value = "Balance Sheet";

                worksheet.Cell("A1").Style.Font.Bold = true;
                worksheet.Cell("A1").Style.Font.FontSize = 20;
                worksheet.Cell("A2").Value = "Balance Sheet";
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("B2").Value = "From : " + pBalanceSheet.FromDate.ToShortDateString() + "";
                worksheet.Cell("C2").Value = "To : " + pBalanceSheet.ToDate.ToShortDateString() + "";
                worksheet.Cell("E2").Style.Font.Bold = true;

                worksheet.Cell("A3").Value = "";
                worksheet.Cell("C3").Value = "                                              " + SysPrefs.DefaultCurrency;
                worksheet.Cell("C3").Style.Font.Bold = true;
                worksheet.Cell("A3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("B3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("C3").Style.Fill.BackgroundColor = XLColor.PersianRed;

                worksheet.Cell("A3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("B3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("C3").Style.Font.FontColor = XLColor.White;

                worksheet.ColumnWidth = 25;
                int mycount1 = 4;
                worksheet.Cell("A" + mycount1).Value = "Non Current Assets";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;

                #region Non Current Assets 
                string sqlNonCurrentAssets = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.BSNonCurrentAssetsAccount + "' UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
                DataTable dtNonCurrentAssets = DBUtils.GetDataTable(sqlNonCurrentAssets);
                if (dtNonCurrentAssets != null)
                {
                    if (dtNonCurrentAssets.Rows.Count > 0)
                    {
                        TotalTangibleAssets = GetBalance(Common.toString(dtNonCurrentAssets.Rows[0]["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate); ;
                    }
                    foreach (DataRow drAccount in dtNonCurrentAssets.Rows)
                    {
                        bool isHead = Common.toBool(drAccount["isHead"]);
                        spaces = "";
                        level = Common.toInt(drAccount["level"]);
                        string color = "PersianRed";
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }

                        decimal Currbalance = GetBalance(Common.toString(drAccount["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                        if (Common.toDecimal(Currbalance) == 0)
                        {
                            if (isHead)
                            {
                                if (pBalanceSheet.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Currbalance;

                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }

                        }
                        else
                        {
                            if (Common.toDecimal(Currbalance) > 0)
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = "(" + Math.Abs(Currbalance) + ")";
                            }

                        }
                        if (level == 0 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                        }
                        else if (level == 1 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                        }
                        else if (level == 2 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                        }
                        else if (level == 3 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                        }
                        //else
                        //{
                        //    worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //}
                    }

                }
                #endregion
                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total Tangible Assets";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalTangibleAssets;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 2;
                worksheet.Cell("A" + mycount1).Value = "Current Assets";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;
                #region  Current Assets 
                pBalanceSheet.AccountCode = SysPrefs.BSCurrentAssetsAccount;
                string sqlCurrentAssets = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.BSCurrentAssetsAccount + "' UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
                DataTable dtCurrentAssets = DBUtils.GetDataTable(sqlCurrentAssets);
                if (dtCurrentAssets != null)
                {
                    if (dtCurrentAssets.Rows.Count > 0)
                    {
                        TotalCurrentAssets = GetBalance(Common.toString(pBalanceSheet.AccountCode), pBalanceSheet.FromDate, pBalanceSheet.ToDate); ;
                    }
                    foreach (DataRow drAccount in dtCurrentAssets.Rows)
                    {
                        bool isHead = Common.toBool(drAccount["isHead"]);
                        spaces = "";
                        level = Common.toInt(drAccount["level"]);
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }
                        decimal Currbalance = GetBalance(Common.toString(drAccount["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                        if (Common.toDecimal(Currbalance) == 0)
                        {
                            if (isHead)
                            {
                                if (pBalanceSheet.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Currbalance;
                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }
                        }
                        else
                        {
                            if (Common.toDecimal(Currbalance) > 0)
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = "(" + Math.Abs(Currbalance) + ")";
                            }
                        }
                        if (level == 0 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                        }
                        else if (level == 1 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                        }
                        else if (level == 2 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                        }
                        else if (level == 3 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                        }
                        //else
                        //{
                        //    worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //}
                    }
                }
                #endregion

                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total Current Assets";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalCurrentAssets;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 2;

                TotalAssets = TotalTangibleAssets + TotalCurrentAssets;

                worksheet.Cell("A" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("A" + mycount1).Value = "Total Assets";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("B" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("C" + mycount1).Value = TotalAssets;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 3;
                worksheet.Cell("A" + mycount1).Value = "Creditors: amounts falling due within one year";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;

                #region  Current Liabilities 
                string sqlCurrentLiabilities = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.BSCurrentLiabilitesAccount + "' UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
                DataTable dtCurrentLiabilities = DBUtils.GetDataTable(sqlCurrentLiabilities);
                if (dtCurrentLiabilities != null)
                {
                    if (dtCurrentLiabilities.Rows.Count > 0)
                    {
                        TotalCurrentLiabilities = GetBalance(Common.toString(dtCurrentLiabilities.Rows[0]["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                    }

                    foreach (DataRow drAccount in dtCurrentLiabilities.Rows)
                    {
                        bool isHead = Common.toBool(drAccount["isHead"]);
                        spaces = "";
                        level = Common.toInt(drAccount["level"]);
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }
                        decimal Currbalance = GetBalance(Common.toString(drAccount["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                        if (Common.toDecimal(Currbalance) == 0)
                        {
                            if (isHead)
                            {
                                if (pBalanceSheet.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Currbalance;
                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }
                        }
                        else
                        {
                            if (Common.toDecimal(Currbalance) > 0)
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = "(" + Math.Abs(Currbalance) + ")";
                            }
                        }
                        if (level == 0)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                        }
                        else if (level == 1)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                        }
                        else if (level == 2 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                        }
                        else if (level == 3 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                        }
                        //else
                        //{
                        //    worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //}
                    }
                }
                #endregion
                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total Creditors: amounts falling due within one year";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalCurrentLiabilities;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 2;

                TotalAssetsLessCurrentLiabilities = TotalTangibleAssets + (TotalAssets - TotalCurrentLiabilities);

                worksheet.Cell("A" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("A" + mycount1).Value = "Total Assets less Current Liabilities";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("B" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("C" + mycount1).Value = TotalAssetsLessCurrentLiabilities;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 3;

                worksheet.Cell("A" + mycount1).Value = "Creditors: amounts falling due after more than one year";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;

                #region  Non Current Liabilities 
                string sqlNonCurrentLiabilities = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.BSNonCurrentliabilitesAccount + "' UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
                DataTable dtNonCurrentLiabilities = DBUtils.GetDataTable(sqlNonCurrentLiabilities);
                if (dtNonCurrentLiabilities != null)
                {
                    if (dtNonCurrentLiabilities.Rows.Count > 0)
                    {
                        TotalNonCurrentLiabilities = GetBalance(Common.toString(dtNonCurrentLiabilities.Rows[0]["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                    }
                    foreach (DataRow drAccount in dtNonCurrentLiabilities.Rows)
                    {
                        bool isHead = Common.toBool(drAccount["isHead"]);
                        spaces = "";
                        level = Common.toInt(drAccount["level"]);
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }
                        decimal Currbalance = GetBalance(Common.toString(drAccount["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                        if (Common.toDecimal(Currbalance) == 0)
                        {
                            if (isHead)
                            {
                                if (pBalanceSheet.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Currbalance;
                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }

                        }
                        else
                        {
                            if (Common.toDecimal(Currbalance) > 0)
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = "(" + Math.Abs(Currbalance) + ")";
                            }
                        }
                        if (level == 0 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                        }
                        else if (level == 1 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                        }
                        else if (level == 2 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                        }
                        else if (level == 3 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                        }
                        //else
                        //{
                        //    worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //}
                    }
                }
                #endregion
                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total Creditors: amounts falling due after more than one year";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalNonCurrentLiabilities;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                mycount1 = mycount1 + 3;

                worksheet.Cell("A" + mycount1).Value = "Capital and Reserves";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("A" + mycount1).Style.Font.FontSize = 15;

                #region  Capital and Reserves
                string sqlCapitalandReserves = ";WITH ret AS (SELECT  AccountCode, AccountName, isHead, 0 as level FROM GLChartOfAccounts WHERE AccountCode = '" + SysPrefs.BSCapitalNReservesAccount + "' UNION ALL SELECT  t.AccountCode, t.AccountName, t.isHead, r.level + 1 FROM GLChartOfAccounts t INNER JOIN ret r ON t.ParentAccountCode = r.AccountCode) SELECT AccountCode, AccountName, isHead, level FROM ret order by AccountCode , level";
                DataTable dtCapitalandReserves = DBUtils.GetDataTable(sqlCapitalandReserves);
                if (dtCapitalandReserves != null)
                {
                    if (dtCapitalandReserves.Rows.Count > 0)
                    {
                        TotalCapitalandReserves = GetBalance(Common.toString(dtCapitalandReserves.Rows[0]["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                    }

                    foreach (DataRow drAccount in dtCapitalandReserves.Rows)
                    {
                        bool isHead = Common.toBool(drAccount["isHead"]);
                        spaces = "";
                        level = Common.toInt(drAccount["level"]);
                        if (level == 1)
                        {
                            spaces = "   ";
                        }
                        else if (level == 2)
                        {
                            spaces = "      ";
                        }
                        else if (level == 3)
                        {
                            spaces = "         ";
                        }
                        else if (level == 4)
                        {
                            spaces = "            ";
                        }
                        else if (level == 5)
                        {
                            spaces = "               ";
                        }
                        decimal Currbalance = GetBalance(Common.toString(drAccount["AccountCode"]), pBalanceSheet.FromDate, pBalanceSheet.ToDate);
                        if (Common.toDecimal(Currbalance) == 0)
                        {
                            if (isHead)
                            {
                                if (pBalanceSheet.ShowAll)
                                {
                                    mycount1++;
                                    worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                    worksheet.Cell("C" + mycount1).Value = Currbalance;
                                }
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }

                        }
                        else
                        {
                            if (Common.toDecimal(Currbalance) > 0)
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = Currbalance;
                            }
                            else
                            {
                                mycount1++;
                                worksheet.Cell("A" + mycount1).Value = spaces + drAccount["AccountName"].ToString();
                                worksheet.Cell("C" + mycount1).Value = "(" + Math.Abs(Currbalance) + ")";
                            }
                        }
                        if (level == 0 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Alizarin;
                        }
                        else if (level == 1 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                        }
                        else if (level == 2 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Aqua;
                        }
                        else if (level == 3 && !isHead)
                        {
                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Almond;
                        }
                        //else
                        //{
                        //    worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //    worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.Amaranth;
                        //}
                    }
                }
                #endregion

                mycount1++;
                worksheet.Cell("A" + mycount1).Value = "Total Capital and Reserves";
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Value = TotalCapitalandReserves;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;

                TotalCapitalandLiabilites = TotalNonCurrentLiabilities + TotalCapitalandReserves;
                mycount1 = mycount1 + 2;
                worksheet.Cell("A" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("A" + mycount1).Value = "Total Capital and Liabilites";
                worksheet.Cell("A" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("B" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Style.Border.TopBorder = 0;
                worksheet.Cell("C" + mycount1).Value = TotalCapitalandLiabilites;
                worksheet.Cell("C" + mycount1).Style.Border.BottomBorder = 0;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;

                worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.White;

                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("B" + mycount1).Style.Font.Bold = true;
                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                string HostName = System.Web.HttpContext.Current.Request.Url.Host;
                string fileName = Common.getRandomDigit(4) + "-" + "BalanceSheet";
                workbook.SaveAs(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".xlsx"));
                Response.Clear();
                Response.ClearHeaders();
                Response.ClearContent();
                Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".xlsx");
                Response.ContentType = "text/plain";
                Response.Flush();
                //Response.TransmitFile(Server.MapPath("~/Uploads/" + "/1/DeanRpt.xlsx"));
                Response.TransmitFile(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/" + fileName + ".xlsx"));
                Response.End();

            }
            catch (Exception x)
            {
                Response.Write("<font color='Red'>" + x.Message + " </font>");
            }

        }
        #endregion

        #region Buyer Report
        public ActionResult BuyerReport()
        {
            BuyerReportViewModel pBalanceSheet = new BuyerReportViewModel();
            pBalanceSheet.FromDate = Common.toDateTime(SysPrefs.PostingDate).AddDays(-30); ;
            pBalanceSheet.ToDate = Common.toDateTime(SysPrefs.PostingDate);
            pBalanceSheet.isvalid = false;
            pBalanceSheet.ShowAll = false;
            return View("~/Views/Reports/Accounts/BuyerReport.cshtml", pBalanceSheet);
        }
        public ActionResult BuyerReport2(BuyerReportViewModel pTrialbalanceViewModel)

        {
            //  pTrialbalanceViewModel = TrailBalanceTableCopy(pTrialbalanceViewModel);
            return View("~/Views/Reports/Accounts/BuyerReport2.cshtml", pTrialbalanceViewModel);
        }
        [HttpPost]
        public ActionResult BuyerReport(BuyerReportViewModel pModel)
        {
            if (!string.IsNullOrEmpty(pModel.AccountCode))
            {
                pModel = BuyerReportTable(pModel);
            }
            else
            {
                pModel.ErrMessage = Common.GetAlertMessage(1, "Please select account code.");
            }

            return View("~/Views/Reports/Accounts/BuyerReport.cshtml", pModel);
        }
        public static BuyerReportViewModel BuyerReportTable(BuyerReportViewModel pReport)
        {
            pReport.isvalid = true;
            StringBuilder strMyTable = new StringBuilder();
            #region Header
            string AccountName = PostingUtils.getGLAccountNameByCode(pReport.AccountCode);
            if (pReport.IsPDF)
            {
                strMyTable.Append("<span style='text-align=center;'><h3>Buyer Report </h3><span>");
                // strMyTable.Append("<img src='http://localhost:42625/assets/img/iremitfylogo.png' alt='' />");
                strMyTable.Append("<br>");
            }
            strMyTable.Append("<b>Account: " + AccountName + "  (" + pReport.AccountCode + ") &nbsp; &nbsp; &nbsp; &nbsp ");
            strMyTable.Append("  From Date: " + pReport.FromDate.ToShortDateString() + " &nbsp; &nbsp; &nbsp; &nbsp; To Date: " + pReport.ToDate.ToShortDateString() + "</b>");
            strMyTable.Append("<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>");
            strMyTable.Append("<table style='table-layout: auto; empty-cells: show;' border='1' rules='rows'>");
            strMyTable.Append("<thead>");
            strMyTable.Append("<tr>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width: 10%;text-align:center;'>Date</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width: 10%;text-align:center;'># of Trans</th>");
            strMyTable.Append("<th class='tabelHeader' colspan='4' style='border: 1px solid antiquewhite; width: 50%;text-align:center;'>Fund Transfers</th>");
            strMyTable.Append("<th class='tabelHeader' colspan='2' style='border: 1px solid antiquewhite; text-align:center;width: 30%;'>Balance</th>");
            strMyTable.Append("</tr>");
            strMyTable.Append("<tr>");
            strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' align='center'></th>");
            strMyTable.Append("<th class='tabelHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>Debit FC</th>");
            strMyTable.Append("<th class='tabelHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>Debit LC</th>");
            strMyTable.Append("<th class='tabelHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>Credit FC </th>");
            strMyTable.Append("<th class='tabelHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>Credit LC </th>");
            strMyTable.Append("<th class='tabelHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>FC Amount</th>");
            strMyTable.Append("<th class='tabelHeader' scope='col' style='text-align:right; border: 1px solid antiquewhite;'>LC Amount</th>");
            strMyTable.Append("</tr>");
            strMyTable.Append("</thead>");
            strMyTable.Append("<tbody>");
            #endregion Header

            decimal TotalFCDebit = 0.0m;
            decimal TotalFCCredit = 0.0m;
            decimal TotalLCDebit = 0.0m;
            decimal TotalLCCredit = 0.0m;
            decimal TotalFCBalance = 0.0m;
            decimal TotalLCBalance = 0.0m;

            #region Opening balance

            AccountBalance accountBalance = new AccountBalance();
            accountBalance = PostingUtils.GetOpeningBalance(pReport.AccountCode, pReport.FromDate);
            TotalFCBalance = Common.toDecimal(accountBalance.ForeignCurrencyBalance);
            TotalLCBalance = Common.toDecimal(accountBalance.LocalCurrencyBalance);

            strMyTable.Append("<tr>");
            strMyTable.Append("<td align='left'> " + pReport.FromDate.AddDays(-1).ToShortDateString() + "</td>");
            strMyTable.Append("<td align='left'>Opening Balance</td>");
            strMyTable.Append("<td colspan='2' align='left'></td>");
            strMyTable.Append("<td colspan='2' align='left'></td>");
            strMyTable.Append("<td align='right'> " + Math.Round(accountBalance.ForeignCurrencyBalance, 2) + "</td>");
            strMyTable.Append("<td align='right'> " + Math.Round(accountBalance.LocalCurrencyBalance, 2) + "</td>");
            strMyTable.Append("</tr>");
            #endregion

            string mySql = "select AddedDate, count(*) as Total, ";
            mySql += "sum(case when LocalCurrencyAmount > 0 then LocalCurrencyAmount else 0 end) as DebitLCTotal, ";
            mySql += "sum(case when LocalCurrencyAmount <= 0 then LocalCurrencyAmount else 0 end) as CreditLCTotal,  ";
            mySql += "sum(case when ForeignCurrencyAmount > 0 then ForeignCurrencyAmount else 0 end) as DebitFCTotal,  ";
            mySql += "sum(case when ForeignCurrencyAmount <= 0 then ForeignCurrencyAmount else 0 end) as CreditFCTotal ";
            mySql += " from GLTransactions where AccountCode = '" + pReport.AccountCode + "' and AddedDate  >= '" + Common.toDateTime(pReport.FromDate).ToString("yyyy-MM-dd 00:00:00") + "' and AddedDate  <= '" + Common.toDateTime(pReport.ToDate).ToString("yyyy-MM-dd 23:59:00") + "' group by AddedDate order by AddedDate";

            DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
            if (dtGLTransactions != null)
            {
                foreach (DataRow MyRow in dtGLTransactions.Rows)
                {
                    TotalFCDebit = TotalFCDebit + Common.toDecimal(MyRow["DebitFCTotal"]);
                    TotalLCDebit = TotalLCDebit + Common.toDecimal(MyRow["DebitLCTotal"]);
                    TotalFCCredit = TotalFCCredit + Common.toDecimal(MyRow["CreditFCTotal"]);
                    TotalLCCredit = TotalLCCredit + Common.toDecimal(MyRow["CreditLCTotal"]);

                    TotalFCBalance = TotalFCBalance + Common.toDecimal(MyRow["DebitFCTotal"]) + Common.toDecimal(MyRow["CreditFCTotal"]);
                    TotalLCBalance = TotalLCBalance + Common.toDecimal(MyRow["DebitLCTotal"]) + Common.toDecimal(MyRow["CreditLCTotal"]);

                    strMyTable.Append("<tr>");
                    strMyTable.Append("<td align='left' > " + Common.toDateTime(MyRow["AddedDate"]).ToShortDateString() + "</td>");
                    strMyTable.Append("<td align='left' ><a href =\"javascript: showDetails('/RptAccounts/ReportView?pCode=" + pReport.AccountCode + "&Date=" + Common.toDateTime(MyRow["AddedDate"]).ToShortDateString() + "','Voucher View')\")>" + MyRow["Total"].ToString() + "</a></td>");
                    // strMyTable.Append("<td align='left' > <a href ='/RptAccounts/ReportView?pCode=" + pReport.AccountCode + "&Date=" + Common.toDateTime(MyRow["AddedDate"]).ToShortDateString() + "'>" + MyRow["Total"].ToString() + "</a></td>");
                    strMyTable.Append("<td align='right' > " + Math.Round(Common.toDouble(MyRow["DebitFCTotal"]), 2) + " </td>");
                    strMyTable.Append("<td align='right'  >" + Math.Round(Common.toDouble(MyRow["DebitLCTotal"]), 2) + "</td>");
                    strMyTable.Append("<td align='right' > " + Math.Round(Common.toDouble(MyRow["CreditFCTotal"]), 2) + " </td>");
                    strMyTable.Append("<td align='right'  >" + Math.Round(Common.toDouble(MyRow["CreditLCTotal"]), 2) + "</td>");
                    strMyTable.Append("<td align='right' > " + Math.Round(TotalFCBalance, 2) + " </td>");
                    strMyTable.Append("<td align='right'  >" + Math.Round(TotalLCBalance, 2) + "</td>");
                    strMyTable.Append("</tr>");
                }
                strMyTable.Append("<tr>");
                strMyTable.Append("<td colspan='2' style='background-color:rgb(255, 188, 0)'><b>Total</b></td>");
                strMyTable.Append("<td style='text-align:right; background-color:rgb(255, 188, 0)'>" + Math.Round(TotalFCDebit, 2) + "</td>");
                strMyTable.Append("<td style='text-align:right; background-color:rgb(255, 188, 0)'>" + Math.Round(TotalLCDebit, 2) + "</td>");
                strMyTable.Append("<td style='text-align:right; background-color:rgb(255, 188, 0)'>" + Math.Round(TotalFCCredit, 2) + "</td>");
                strMyTable.Append("<td style='text-align:right; background-color:rgb(255, 188, 0)'>" + Math.Round(TotalLCCredit, 2) + "</td>");
                strMyTable.Append("<td style='text-align:right; background-color:rgb(255, 188, 0)'></td>");
                strMyTable.Append("<td style='text-align:right; background-color:rgb(255, 188, 0)'></td>");
                strMyTable.Append("</tr>");
            }
            strMyTable.Append("</tbody></table></div>");
            pReport.HtmlTable = strMyTable.ToString();
            return pReport;
        }
        public ActionResult BuyerReport2_Read([DataSourceRequest] DataSourceRequest request)
        {
            #region Get Filter Values

            string ParamAccountCode = "";
            string ParamFromDate = "";
            string ParamToDate = "";

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
                            if (string.IsNullOrEmpty(ParamFromDate))
                            {
                                ParamFromDate = Common.toString(descriptor.Value);
                            }
                        }
                        if (descriptor.Member == "AddedDate")
                        {
                            ParamToDate = Common.toString(descriptor.Value);
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
                                if (descriptor2.Member == "AddedDate")
                                {
                                    if (string.IsNullOrEmpty(ParamFromDate))
                                    {
                                        ParamFromDate = Common.toString(descriptor2.Value);
                                    }
                                }
                                if (descriptor2.Member == "AddedDate")
                                {
                                    ParamToDate = Common.toString(descriptor2.Value);
                                }

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

                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            decimal TotalFCBalance = 0.0m;
            decimal TotalLCBalance = 0.0m;
            #region old code
            if (!string.IsNullOrEmpty(ParamAccountCode))
            {
                decimal TotalFCDebit = 0.0m;
                decimal TotalFCCredit = 0.0m;
                decimal TotalLCDebit = 0.0m;
                decimal TotalLCCredit = 0.0m;

                #region Opening balance

                AccountBalance accountBalance = new AccountBalance();
                accountBalance = PostingUtils.GetOpeningBalance(ParamAccountCode, Common.toDateTime(ParamFromDate));
                TotalFCBalance = Common.toDecimal(accountBalance.ForeignCurrencyBalance);
                TotalLCBalance = Common.toDecimal(accountBalance.LocalCurrencyBalance);

                #endregion

                const string countQuery = @"SELECT COUNT(1) FROM GLTransactions  /**where**/";
                string selectQuery = @"SELECT  * ";
                selectQuery += @"  AddedDate, count(*) as Total, ForeignCurrencyAmount, LocalCurrencyAmount,sum(case when LocalCurrencyAmount > 0 then LocalCurrencyAmount else 0 end) as DebitLCTotal,sum(case when LocalCurrencyAmount <= 0 then LocalCurrencyAmount else 0 end) as CreditLCTotal,sum(case when ForeignCurrencyAmount > 0 then ForeignCurrencyAmount else 0 end) as DebitFCTotal,sum(case when ForeignCurrencyAmount <= 0 then ForeignCurrencyAmount else 0 end) as CreditFCTotal  from GLTransactions ";
                selectQuery += @" /**where**/ ) AS RowConstrainedResult WHERE   RowNum >= (@PageIndex * @PageSize + 1 ) AND RowNum <= (@PageIndex + 1) * @PageSize ORDER BY RowNum";

                SqlBuilder builder = new SqlBuilder();
                var count = builder.AddTemplate(countQuery);

                var selector = builder.AddTemplate(selectQuery);
                builder.Where(" AccountCode ='" + ParamAccountCode + "'  and AddedDate  >= '" + Common.toDateTime(ParamFromDate).ToString("yyyy-MM-dd 00:00:00") + "' and AddedDate  <= '" + Common.toDateTime(ParamToDate).ToString("yyyy-MM-dd 23:59:00") + "'  group by AddedDate");
                if (request.Sorts != null && request.Sorts.Any())
                {
                    builder = Common.ApplySorting(builder, request.Sorts);
                }
                else
                {
                    builder.OrderBy("AddedDate asc");
                }

                var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);
                var rows = _dbcontext.Query<BuyerReportViewModel>(selector.RawSql, selector.Parameters);

                List<BuyerReportViewModel> objList = new List<BuyerReportViewModel>();
                foreach (var item in rows)
                {
                    TotalFCDebit += Common.toDecimal(item.DebitFCTotal);
                    TotalLCDebit += Common.toDecimal(item.DebitLCTotal);
                    TotalFCCredit += Common.toDecimal(item.CreditFCTotal);
                    TotalLCCredit += Common.toDecimal(item.CreditLCTotal);

                    TotalFCBalance = TotalFCBalance + Common.toDecimal(item.DebitFCTotal) + Common.toDecimal(item.CreditFCTotal);
                    TotalLCBalance = TotalLCBalance + Common.toDecimal(item.DebitLCTotal) + Common.toDecimal(item.CreditLCTotal);

                    objList.Add(new BuyerReportViewModel
                    {
                        //      strMyTable.Append("<td align='left' > " + Common.toDateTime(MyRow["AddedDate"]).ToShortDateString() + "</td>");
                        //<a href =\"javascript: showDetails('/RptAccounts/ReportView?pCode=" + pReport.AccountCode + "&Date=" + Common.toDateTime(MyRow["AddedDate"]).ToShortDateString() + "','Voucher View')\")>" + MyRow["Total"].ToString() + "</a></td>");
                        ////  <a href ='/RptAccounts/ReportView?pCode=" + pReport.AccountCode + "&Date=" + Common.toDateTime(MyRow["AddedDate"]).ToShortDateString() + "'>" + MyRow["Total"].ToString() + "</a></td>");
                        // Math.Round(Common.toDouble(MyRow["DebitFCTotal"]), 2) + " </td>");
                        // Math.Round(Common.toDouble(MyRow["DebitLCTotal"]), 2) + "</td>");
                        // Math.Round(Common.toDouble(MyRow["CreditFCTotal"]), 2) + " </td>");
                        // Math.Round(Common.toDouble(MyRow["CreditLCTotal"]), 2) + "</td>");
                        // Math.Round(TotalFCBalance, 2) + " </td>");
                        // Math.Round(TotalLCBalance, 2) + "</td>");

                        pDate = Common.toDateTime(item.AddedDate).ToShortDateString(),
                        pTotal = "<a href =\"javascript: showDetails('/RptAccounts/ReportView?pCode=" + ParamAccountCode + "&Date=" + Common.toDateTime(item.AddedDate).ToShortDateString() + "','Voucher View')\")>" + item.Total.ToString() + "</a>",
                        DebitFCTotal = Math.Round(Common.toDecimal(item.DebitFCTotal), 2),
                        DebitLCTotal = Math.Round(Common.toDecimal(item.DebitLCTotal), 2),
                        CreditFCTotal = Math.Round(Common.toDecimal(item.CreditFCTotal), 2),
                        CreditLCTotal = Math.Round(Common.toDecimal(item.CreditLCTotal), 2)

                    });
                }
                #region Opening Balance  
                List<BuyerReportViewModel> openingBalance = new List<BuyerReportViewModel>();

                if (!string.IsNullOrEmpty(Common.toString(ParamAccountCode)))
                {
                    //AccountBalance AcBalances = PostingUtils.GetOpeningBalance(ParamAccountCode, Common.toDateTime(ParamFromDate));
                    //decimal LocalCurrencyBalance = Common.toDecimal(AcBalances.LocalCurrencyBalance);
                    //decimal ForeignCurrencyBalance = Common.toDecimal(AcBalances.ForeignCurrencyBalance);
                    if (!string.IsNullOrEmpty(Common.toString(ParamAccountCode)))
                    {
                        openingBalance.Add(new BuyerReportViewModel
                        {
                            VoucherNumber = "Opening Balance - " + Common.toDateTime(ParamFromDate).ToShortDateString(),
                            LocalCurrencyBalance = Math.Round(accountBalance.LocalCurrencyBalance, 2),
                            ForeignCurrencyBalance = Math.Round(accountBalance.ForeignCurrencyBalance, 2),
                            SortOrder = 1,
                        });
                    }
                }

                #endregion

                var result = new DataSourceResult()
                {
                    Data = objList,
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

            #endregion

            #region
            //if (request.PageSize == 0)
            //{
            //    request.PageSize = 50;
            //}

            //string sql = @"SELECT    AddedDate, count(*) as Total,sum(case when LocalCurrencyAmount > 0 then LocalCurrencyAmount else 0 end) as DebitLCTotal,sum(case when LocalCurrencyAmount <= 0 then LocalCurrencyAmount else 0 end) as CreditLCTotal,sum(case when ForeignCurrencyAmount > 0 then ForeignCurrencyAmount else 0 end) as DebitFCTotal,sum(case when ForeignCurrencyAmount <= 0 then ForeignCurrencyAmount else 0 end) as CreditFCTotal  from GLTransactions  WHERE  AccountCode ='"+ ParamAccountCode + "' and AddedDate  >= '" + Common.toDateTime(ParamFromDate).ToString("yyyy-MM-dd 00:00:00") + "' and AddedDate  <= '" + Common.toDateTime(ParamToDate).ToString("yyyy-MM-dd 23:59:00") + "'  group by AddedDate ";
            //var objList = _dbcontext.Query<BuyerReportViewModel>(sql);
            ////List<BuyerReportViewModel> objList = new List<BuyerReportViewModel>();
            ////foreach (var item in data)
            ////{
            ////    objList.Add(new BuyerReportViewModel
            ////    {
            ////        AiuUserID = item.AiuUserID,
            ////        RepDate = item.RepDate,
            ////        FirstName = item.FirstName,
            ////        LastName = item.LastName,
            ////        Country = item.Country,
            ////        UserName = item.UserName,
            ////        Email = item.Email,
            ////        HomePhone = item.HomePhone,
            ////        CellPhone = item.CellPhone,
            ////        OfficePhone = item.OfficePhone,
            ////        Status = item.Status,
            ////        CriExceedMax = item.CriExceedMax,
            ////        ProspectsPerDay = item.ProspectsPerDay,
            ////        CampusTitle = item.CampusTitle,
            ////        RepName = item.RepName,
            ////        RepExt = item.RepExt

            ////    });

            ////}
            //var total = objList.Count();
            //if (request.Page > 0)
            //{
            //    objList = objList.Skip((request.Page - 1) * request.PageSize).ToList();
            //}
            //objList = objList.Take(request.PageSize).ToList();
            //var result = new DataSourceResult()
            //{
            //    Data = objList,
            //    Total = total
            //};
            //return Json(result);
            #endregion
        }
        //public void BuyerReportExportToPDF(BuyerReportViewModel pReport)
        //{
        //    pReport.IsPDF = true;
        //    pReport = BuyerReportTable(pReport);
        //    string fileName = Common.getRandomDigit(4) + "-" + "BuyerReport";
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/");
        //    HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //    htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //    PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(pReport.HtmlTable, sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    Response.Clear();
        //    Response.ClearContent();
        //    Response.ClearHeaders();
        //    Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".pdf");
        //    Response.ContentType = "application/pdf";
        //    Response.Flush();
        //    Response.TransmitFile(Server.MapPath("~/Uploads/" + fileName + ".pdf"));
        //    Response.End();
        //}
        public void BuyerReportExportToExcel(BuyerReportViewModel pModel)
        {
            try
            {
                string AccountName = DBUtils.executeSqlGetSingle("select AccountName from GLChartOfAccounts where AccountCode = '" + pModel.AccountCode + "'");
                XLWorkbook workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Buyer Report");
                worksheet.Cell("A1").Value = "Buyer Report";
                worksheet.Cell("A1").Style.Font.Bold = true;

                worksheet.Cell("C1").Value = "Account :" + AccountName + "  (" + pModel.AccountCode + ")";
                worksheet.Cell("C1").Style.Font.Bold = true;
                worksheet.Cell("A2").Value = " From Date:" + Common.toDateTime(pModel.FromDate).ToShortDateString() + "";
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("D2").Value = " To Date: " + Common.toDateTime(pModel.ToDate).ToShortDateString() + "";
                worksheet.Cell("D2").Style.Font.Bold = true;

                worksheet.Cell("A3").Value = "Date";
                worksheet.Cell("B3").Value = "No of Transactions";
                worksheet.Cell("C3").Value = "Debit FC Amount";
                worksheet.Cell("D3").Value = "Debit LC Amount";
                worksheet.Cell("E3").Value = "Credit FC Amount";
                worksheet.Cell("F3").Value = "Credit LC Amount";
                worksheet.Cell("G3").Value = "Balance FC Amount";
                worksheet.Cell("H3").Value = "Balance LC Amount";

                worksheet.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("B3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("C3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("D3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("E3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("F3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("G3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("H3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;


                worksheet.Cell("A3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("B3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("C3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("D3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("E3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("F3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("G3").Style.Fill.BackgroundColor = XLColor.PersianRed;
                worksheet.Cell("H3").Style.Fill.BackgroundColor = XLColor.PersianRed;

                worksheet.Cell("A3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("B3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("C3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("D3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("E3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("F3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("G3").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("H3").Style.Font.FontColor = XLColor.White;
                worksheet.ColumnWidth = 25;
                int mycount1 = 3;
                decimal TotalFCDebit = 0.0m;
                decimal TotalFCCredit = 0.0m;
                decimal TotalLCDebit = 0.0m;
                decimal TotalLCCredit = 0.0m;
                decimal TotalFCBalance = 0.0m;
                decimal TotalLCBalance = 0.0m;

                AccountBalance accountBalance = new AccountBalance();
                accountBalance = PostingUtils.GetOpeningBalance(pModel.AccountCode, pModel.FromDate);
                TotalFCBalance = Common.toDecimal(accountBalance.ForeignCurrencyBalance);
                TotalLCBalance = Common.toDecimal(accountBalance.LocalCurrencyBalance);

                TotalFCBalance = Common.toDecimal(accountBalance.ForeignCurrencyBalance);
                TotalLCBalance = Common.toDecimal(accountBalance.LocalCurrencyBalance);

                mycount1++;
                worksheet.Cell("A" + mycount1).Value = pModel.FromDate.AddDays(-1).ToShortDateString();
                worksheet.Cell("B" + mycount1).Value = "Opening balance";
                worksheet.Cell("C" + mycount1).Value = "";
                worksheet.Cell("D" + mycount1).Value = "";
                worksheet.Cell("E" + mycount1).Value = "";
                worksheet.Cell("F" + mycount1).Value = "";
                worksheet.Cell("G" + mycount1).Value = Math.Round(accountBalance.ForeignCurrencyBalance, 2);
                worksheet.Cell("H" + mycount1).Value = Math.Round(accountBalance.LocalCurrencyBalance, 2);

                worksheet.Cell("A" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("B" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("C" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("D" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("E" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("F" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("G" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("H" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                string mySql = "select AddedDate, count(*) as Total, ";
                mySql += "sum(case when LocalCurrencyAmount > 0 then LocalCurrencyAmount else 0 end) as DebitLCTotal, ";
                mySql += "sum(case when LocalCurrencyAmount <= 0 then LocalCurrencyAmount else 0 end) as CreditLCTotal,  ";
                mySql += "sum(case when ForeignCurrencyAmount > 0 then ForeignCurrencyAmount else 0 end) as DebitFCTotal,  ";
                mySql += "sum(case when ForeignCurrencyAmount <= 0 then ForeignCurrencyAmount else 0 end) as CreditFCTotal ";
                mySql += " from GLTransactions where AccountCode = '" + pModel.AccountCode + "' and AddedDate  >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy-MM-dd 00:00:00") + "' and AddedDate  <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy-MM-dd 23:59:00") + "' group by AddedDate order by AddedDate";

                //mySql = "SELECT GLChartOfAccounts.AccountName, GLChartOfAccounts.AccountCode, sum(GLTransactions.LocalCurrencyAmount) as LocalCurrencyBalance FROM GLChartOfAccounts INNER JOIN GLTransactions ON GLChartOfAccounts.AccountCode = GLTransactions.AccountCode where isHead = 1 group by GLChartOfAccounts.AccountCode, GLChartOfAccounts.AccountName";
                DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
                if (dtGLTransactions != null)
                {
                    foreach (DataRow MyRow in dtGLTransactions.Rows)
                    {
                        TotalFCDebit = TotalFCDebit + Common.toDecimal(MyRow["DebitFCTotal"]);
                        TotalLCDebit = TotalLCDebit + Common.toDecimal(MyRow["DebitLCTotal"]);
                        TotalFCCredit = TotalFCCredit + Common.toDecimal(MyRow["CreditFCTotal"]);
                        TotalLCCredit = TotalLCCredit + Common.toDecimal(MyRow["CreditLCTotal"]);

                        TotalFCBalance = TotalFCBalance + Common.toDecimal(MyRow["DebitFCTotal"]) + Common.toDecimal(MyRow["CreditFCTotal"]);
                        TotalLCBalance = TotalLCBalance + Common.toDecimal(MyRow["DebitLCTotal"]) + Common.toDecimal(MyRow["CreditLCTotal"]);

                        mycount1++;
                        // table.Rows.Add(subName3, subName2, subName1, MyRow["AccountName"].ToString(), strAmount, strDebit, strCredit);
                        worksheet.Cell("A" + mycount1).Value = Math.Round(Common.toDouble(MyRow["AddedDate"]), 2);
                        worksheet.Cell("B" + mycount1).Value = Math.Round(Common.toDouble(MyRow["Total"]), 2);
                        worksheet.Cell("C" + mycount1).Value = Math.Round(Common.toDouble(MyRow["DebitFCTotal"]), 2);
                        worksheet.Cell("D" + mycount1).Value = Math.Round(Common.toDouble(MyRow["DebitLCTotal"]), 2);
                        worksheet.Cell("E" + mycount1).Value = Math.Round(Common.toDouble(MyRow["CreditFCTotal"]), 2);
                        worksheet.Cell("F" + mycount1).Value = Math.Round(Common.toDouble(MyRow["CreditLCTotal"]), 2);
                        worksheet.Cell("G" + mycount1).Value = Math.Round(Common.toDouble(TotalFCBalance), 2);
                        worksheet.Cell("H" + mycount1).Value = Math.Round(Common.toDouble(TotalLCBalance), 2);



                        worksheet.Cell("A" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell("B" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell("C" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        worksheet.Cell("D" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        worksheet.Cell("E" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        worksheet.Cell("F" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        worksheet.Cell("G" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        worksheet.Cell("H" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                        worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("D" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("G" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                        worksheet.Cell("H" + mycount1).Style.Fill.BackgroundColor = XLColor.White;


                        worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("B" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("D" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("E" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("G" + mycount1).Style.Font.Bold = true;
                        worksheet.Cell("H" + mycount1).Style.Font.Bold = true;

                    }
                    mycount1++;
                    worksheet.Cell("A" + mycount1).Value = "Total";
                    worksheet.Cell("C" + mycount1).Value = Math.Round(TotalFCDebit, 2);
                    worksheet.Cell("D" + mycount1).Value = Math.Round(TotalLCDebit, 2);
                    worksheet.Cell("E" + mycount1).Value = "(" + Math.Round(TotalFCCredit, 2) + ")";
                    worksheet.Cell("F" + mycount1).Value = "(" + Math.Round(TotalLCCredit, 2) + ")";

                    worksheet.Cell("C" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell("D" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell("E" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell("F" + mycount1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                    worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                    worksheet.Cell("D" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                    worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
                    worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.White;

                    worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("D" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("E" + mycount1).Style.Font.Bold = true;
                    worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
                }

                string fileName = Common.getRandomDigit(4) + "-" + "BuyerReport";
                workbook.SaveAs(Server.MapPath("~/Uploads/" + fileName + ".xlsx"));
                Response.Clear();
                Response.ClearHeaders();
                Response.ClearContent();
                Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".xlsx");
                Response.ContentType = "text/plain";
                Response.Flush();
                //Response.TransmitFile(Server.MapPath("~/Uploads/" + "/1/DeanRpt.xlsx"));
                Response.TransmitFile(Server.MapPath("~/Uploads/" + fileName + ".xlsx"));
                Response.End();
            }
            catch (Exception x)
            {
                Response.Write("<font color='Red'>" + x.Message + " </font>");
            }

        }
        #endregion

        #region Daily Account Transactions
        public ActionResult ReportView(string pCode, string Date)
        {
            JournalInquiryViewModel pJournalInquiryViewModel = new JournalInquiryViewModel();
            pJournalInquiryViewModel.AccountCode = pCode;
            //  Common.toDateTime(pReport.FromDate).ToString("yyyy-MM-dd 00:00:00") + "' and AddedDate  <= '" + Common.toDateTime(pReport.ToDate).ToString("yyyy-MM-dd 23:59:00")
            pJournalInquiryViewModel.StartDate = Common.toDateTime(Date).ToString("yyyy-MM-dd 00:00:00");
            pJournalInquiryViewModel.EndDate = Common.toDateTime(Date).ToString("yyyy-MM-dd 23:59:00");
            return View("~/Views/Reports/Accounts/DailyAccountTransactions.cshtml", pJournalInquiryViewModel);
        }

        public ActionResult ViewList_Read([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID  Where GLReferences.GLReferenceID in (select distinct GLReferenceID from GLTransactions /**where**/)";
            string selectQuery = @"SELECT  * FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum";
            selectQuery += @",GLTransactions.GLReferenceID, AccountCode,  (Select AccountName From GLChartOfAccounts Where GLTransactions.AccountCode = GLChartOfAccounts.AccountCode)";
            selectQuery += @"As AccountTitle, GLTransactions.AddedDate as AddedDate, GLTransactions.ForeignCurrencyAmount, GLTransactions.LocalCurrencyAmount, GLTransactions.Memo, ";
            selectQuery += @"GLReferences.VoucherNumber, GLTransactions.TransactionNumber As TransNo, GLTransactions.ForeignCurrencyISOCode, (select FirstName + ' ' + LastName from [Users]";
            selectQuery += @"where [Users].UserID= GLReferences.PostedBy) as PostedBy, (select FirstName + ' ' + LastName from [Users] where [Users].UserID= GLReferences.AuthorizedBy) as AuthBy";
            selectQuery += @",(select Description from ConfigItemsData where GLReferences.TypeiCode= ConfigItemsData.ItemsDataCode) as TypeiCode FROM GLTransactions INNER JOIN GLReferences";
            selectQuery += @" ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID";
            selectQuery += @" Where GLReferences.GLReferenceID in (select distinct GLReferenceID from GLTransactions ";
            selectQuery += @" /**where**/ )) AS RowConstrainedResult WHERE   RowNum >= (@PageIndex * @PageSize + 1 ) AND RowNum <= (@PageIndex + 1) * @PageSize ORDER BY RowNum";

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
            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("VoucherNumber Desc");
            }
            //GLReferenceID in (select GLReferenceID from 

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);
            var rows = _dbcontext.Query<JournalInquiryViewModel>(selector.RawSql, selector.Parameters);
            string voucherNo = "";
            foreach (JournalInquiryViewModel model in rows)
            {
                if (voucherNo != model.VoucherNumber)
                {
                    voucherNo = model.VoucherNumber;
                    model.GLReferenceID = Security.EncryptQueryString(model.GLReferenceID).ToString();
                    model.VoucherNumber = model.VoucherNumber;
                    model.AddedDate = Common.toDateTime(model.AddedDate).ToShortDateString();
                    model.Memo = model.Memo;
                }
                else
                {
                    model.VoucherNumber = "";
                    model.AddedDate = "";
                    model.Memo = "";
                }
            }
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }



        #endregion
    }
}

