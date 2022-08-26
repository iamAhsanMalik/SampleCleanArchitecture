using ClosedXML.Excel;
using DHAAccounts.Models;

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Accounting.Controllers
{
    public class LedgersController : Controller
    {
        // GET: Ledgers
        public ActionResult Index()
        {
            return View();
        }
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
}