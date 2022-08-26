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
    public class RptFinancialsController : Controller
    {
        // GET: RptFinancials
        public ActionResult Index()
        {
            return View();
        }

        #region Exchange Rate History
        public ActionResult ExchangeRateHistory()
        {
            ExchangeRateHistoryViewModel model = new ExchangeRateHistoryViewModel();
            model.FromDate = Common.toDateTime(SysPrefs.PostingDate).AddDays(-30);
            model.ToDate = Common.toDateTime(SysPrefs.PostingDate);
            return View("~/Views/Reports/Financials/ExchangeRateHistory.cshtml", model);
        }
        [HttpPost]
        public ActionResult ExchangeRateHistory(ExchangeRateHistoryViewModel model)
        {
            if (!string.IsNullOrEmpty(model.CurrencyISOCode))
            {
                StringBuilder strMyTable = new StringBuilder();
                strMyTable.Append(ExchangeRateHistoryHeader());
                string mySql = "Select CurrencyISOCode, ExchangeRate,(Select FirstName+' '+LastName from Users where Users.UserID=ExchangeRates.AddedBy) as AddedBy,(Select FirstName+' '+LastName from Users where Users.UserID=ExchangeRates.UpdatedBy) as UpdatedBy,AddedDate, UpdatedBy, UpdatedDate from ExchangeRates";
                if (!string.IsNullOrEmpty(model.CurrencyISOCode))
                {
                    mySql += " Where CurrencyISOCode ='" + model.CurrencyISOCode + "'";
                }
                if (model.FromDate != null)
                {
                    mySql += " and AddedDate >='" + model.FromDate.ToString("yyyy-MM-dd") + " 00:00:00' ";
                }
                if (model.ToDate != null)
                {
                    mySql += " and AddedDate <='" + model.ToDate.ToString("yyyy-MM-dd") + " 23:59:59' ";
                }
                mySql += " order by AddedDate Desc";
                DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
                if (dtGLTransactions != null)
                {
                    foreach (DataRow MyRow in dtGLTransactions.Rows)
                    {
                        strMyTable.Append("<tr>");
                        strMyTable.Append("<td align='left'>" + Common.toDateTime(MyRow["AddedDate"]).ToShortDateString() + "</td>");
                        strMyTable.Append("<td align='left'>" + Common.toString(MyRow["CurrencyISOCode"]) + "</td>");
                        strMyTable.Append("<td align='left'>" + Common.toDecimal(MyRow["ExchangeRate"]) + "</td>");
                        strMyTable.Append("<td align='left'>" + Common.toString(MyRow["AddedBy"].ToString()) + "</td>");
                        strMyTable.Append("<td align='left' style='margin-left:3%'>" + Common.toDateTime(MyRow["UpdatedDate"]).ToShortDateString() + "</td>");
                        strMyTable.Append("<td align='left'style='margin-left:3%'>" + Common.toString(MyRow["UpdatedBy"].ToString()) + "</td>");
                        strMyTable.Append("</tr>");
                    }
                }
                strMyTable.Append("</tbody></table></div>");
                model.Table = strMyTable.ToString();
            }
            else
            {
                model.ErrMessage = Common.GetAlertMessage(1, "Please select currency.");
            }
            return View("~/Views/Reports/Financials/ExchangeRateHistory.cshtml", model);
        }
        public string ExchangeRateHistoryHeader()
        {
            string strMyTableHeader = "<div class='my-table-zebra-rounded' id='tblGLInquiry'  style='width: 100%;'> ";
            strMyTableHeader += "<table style='table-layout: auto; empty-cells: show; width: 100%;' border='1' rules='rows' align='center'>";
            strMyTableHeader += "<thead>";
            strMyTableHeader += "<tr>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Date</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Currency</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Rate</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Added By</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Updated Date</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Updated By</th>";
            strMyTableHeader += "</tr>";
            strMyTableHeader += "</thead>";
            strMyTableHeader += "<tbody>";
            return strMyTableHeader;
        }
        public string ExchangeRateHistoryFooter()
        {
            string strGLJournalInquiryFooter = "";
            return strGLJournalInquiryFooter;
        }
        #endregion

        #region Un Athuorized voucher
        public ActionResult UnAuthorizedVouchers()
        {
            UnAuthorizedVouchersViewModel pJournalInquiryViewModel = new UnAuthorizedVouchersViewModel();
            UnAuthorizedVouchers(pJournalInquiryViewModel);
            return View("~/Views/Reports/Financials/UnAuthorizedVouchers.cshtml");
        }
        [HttpPost]
        public ActionResult UnAuthorizedVouchers(UnAuthorizedVouchersViewModel pJournalInquiryViewModel)
        {
            string isPosted = "";
            StringBuilder strMyTable = new StringBuilder();
            strMyTable.Append(UnAuthorizedVouchersHeader());
            string mySql = "Select GLTransactions.GLReferenceID, (Select AccountName From GLChartOfAccounts Where GLTransactions.AccountCode = GLChartOfAccounts.AccountCode) As AccountTitle, GLTransactions.AddedDate, GLTransactions.ForeignCurrencyAmount, GLTransactions.LocalCurrencyAmount, GLReferences.Memo, GLReferences.VoucherNumber, GLTransactions.TransactionNumber As TransNo, GLTransactions.ForeignCurrencyISOCode, (select FirstName + ' ' + LastName from[Users]where[Users].UserID = GLReferences.PostedBy) as PostedBy, (select FirstName + ' ' + LastName from[Users] where [Users].UserID = GLReferences.AuthorizedBy) as AuthBy,(Select isPosted from GLReferences Where GLTransactions.GLReferenceID = GLReferences.GLReferenceID) as isPosted,(select Description from ConfigItemsData where GLReferences.TypeiCode = ConfigItemsData.ItemsDataCode) as TypeiCode FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID Where isPosted = 0";
            mySql += " order by AddedDate Desc";
            DataTable dtGLTransactions = DBUtils.GetDataTable(mySql);
            if (dtGLTransactions != null)
            {
                foreach (DataRow MyRow in dtGLTransactions.Rows)
                {
                    strMyTable.Append("<tr>");
                    strMyTable.Append("<td align='left'>" + Common.toDateTime(MyRow["AddedDate"]).ToString("yyyy-MM-dd") + "</td>");

                    strMyTable.Append("<td align='left'>" + "<a href=\"javascript: showDetails('/Inquiries/GLTransView?Id=" + Security.EncryptQueryString(MyRow["GLReferenceID"].ToString()) + "','Voucher View')\")>" + MyRow["VoucherNumber"].ToString() + "</a></td>");
                    strMyTable.Append("<td align='left'>" + MyRow["Memo"].ToString() + "</td>");
                    strMyTable.Append("<td align='left'>" + MyRow["AccountTitle"].ToString() + "</td>");
                    strMyTable.Append("<td align='left'>" + MyRow["ForeignCurrencyISOCode"].ToString() + "</td>");
                    strMyTable.Append("<td align='left'>" + MyRow["PostedBy"].ToString() + "</td>");
                    strMyTable.Append("<td align='left'>" + MyRow["TypeiCode"].ToString() + "</td>");
                    string abc = MyRow["isPosted"].ToString();
                    if (MyRow["isPosted"].ToString().Trim() == "False")
                    {
                        isPosted = "UnAthorized";
                    }
                    else
                    {
                        isPosted = "Athorized";
                    }
                    strMyTable.Append("<td align='left'>" + isPosted + "</td>");
                    strMyTable.Append("</tr>");
                }
            }
            strMyTable.Append("</tbody></table></div>");
            pJournalInquiryViewModel.Table = strMyTable.ToString();
            return View("~/Views/Reports/Financials/UnAuthorizedVouchers.cshtml", pJournalInquiryViewModel);
            //ltrTableBody.Text = Convert.ToString(strMyTable);
        }
        public string UnAuthorizedVouchersHeader()
        {
            string strMyTableHeader = "<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>";
            strMyTableHeader += "<table style='table-layout: auto; empty-cells: show;' border='1' rules='rows'>";
            strMyTableHeader += "<thead>";
            strMyTableHeader += "<tr>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Date</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>V.No</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Description</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Account Title</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Curr</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>P.By</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Type</th>";
            strMyTableHeader += "<th class='tabelHeader' align='center'>Authorization status</th>";
            strMyTableHeader += "</tr>";
            strMyTableHeader += "</thead>";
            strMyTableHeader += "<tbody>";
            return strMyTableHeader;
        }
        public string UnAuthorizedVouchersFooter()
        {
            string strGLJournalInquiryFooter = "";
            return strGLJournalInquiryFooter;
        }
        #endregion
    }
}