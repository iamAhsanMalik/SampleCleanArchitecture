using DHAAccounts.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Xml;

namespace Accounting.Controllers
{
    public class AutoRunController : Controller
    {
        // GET: AutoRun
        public ActionResult Index()
        {
            //FileHoldReportExportToPDF(); 
            return View();
        }
        public static string getTransactions()
        {
            #region Header
            StringBuilder strMyTable = new StringBuilder();
            strMyTable.Append("<span style='text-align='center';'><h3>Hold File Report </h3><span><br>");
            strMyTable.Append("<div class='my-table-zebra-rounded' id='tblGLInquiry' style='width: 100%;'>");
            strMyTable.Append("<table style='table-layout: auto; empty-cells: show;' border='1' rules='rows'>");
            strMyTable.Append("<thead>");
            strMyTable.Append("<tr>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width: 10%;' align='center'>TransactionNumber</th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width: 10%;' align='center'>PaidGLReferenceID </ th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width: 10%;' align='center'>Status </th>");
            strMyTable.Append("<th class='tabelHeader' style='border: 1px solid antiquewhite; width: 10%;' align='center'>GLReferenceID </th>");
            strMyTable.Append("</tr>");
            strMyTable.Append("</thead>");
            strMyTable.Append("<tbody>");

            #endregion

            string paidRefId = "", Status = "", GLReferenceID = "", TransactionNumber = "";
            //GLTransactions.GLReferenceID, GLTransactions.TransactionID,GLReferences.VoucherNumber
            string sql = " SELECT GLTransactions.TransactionNumber FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID";
            sql += " WHERE AccountCode = '32120002' and isPosted = '1' ";
            DataTable dtTranx = DBUtils.GetDataTable(sql);
            if (dtTranx != null)
            {
                if (dtTranx.Rows.Count > 0)
                {
                    foreach (DataRow dr in dtTranx.Rows)
                    {
                        TransactionNumber = dr["TransactionNumber"].ToString();
                        if (!string.IsNullOrEmpty(TransactionNumber))
                        {
                            string sqlCustomerTranx = " SELECT PaidGLReferenceID,Status FROM CustomerTransactions WHERE PaymentNo='" + TransactionNumber + "'";
                            DataTable dt = DBUtils.GetDataTable(sqlCustomerTranx);
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                paidRefId = dt.Rows[0]["PaidGLReferenceID"].ToString();
                                Status = dt.Rows[0]["Status"].ToString();
                            }
                            if (!string.IsNullOrEmpty(paidRefId))
                            {
                                string sqlRef = " SELECT GLReferenceID FROM GLReferences WHERE GLReferenceID = " + paidRefId + "";
                                GLReferenceID = DBUtils.executeSqlGetSingle(sqlRef);


                                strMyTable.Append("<tr>");
                                strMyTable.Append("<td align='left'> " + TransactionNumber + "</td>");
                                strMyTable.Append("<td align='left'> " + paidRefId + "</td>");
                                strMyTable.Append("<td align='left'> " + Status + "</td>");
                                strMyTable.Append("<td align='left'> " + GLReferenceID + "</td>");
                                strMyTable.Append("</tr>");

                            }
                        }
                    }
                }

            }
            strMyTable.Append("</tbody></table></div>");
            string HtmlTable = strMyTable.ToString();
            return HtmlTable;

        }
        //public void FileHoldReportExportToPDF()
        //{
        //    string HostName = System.Web.HttpContext.Current.Request.Url.Host;
        //    string htmlTable = getTransactions();
        //    string fileName = Common.getRandomDigit(4) + "-" + "HoldFileReport";
        //    string sPathToWritePdfTo = Server.MapPath("~/Uploads/"+ SysPrefs.SubmissionFolder + "/");
        //    HtmlToPdf htmlToPdfConverter = new HtmlToPdf();
        //    htmlToPdfConverter.SerialNumber = "WBAxCQg8-PhQxOio5-KiFudmh4-aXhpeGBr-YXhraXZp-anZhYWFh";
        //    PdfDocument pdf = htmlToPdfConverter.ConvertHtmlToPdfDocument(htmlTable, sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    pdf.WriteToFile(sPathToWritePdfTo + "\\" + fileName + ".pdf");
        //    Response.Clear();
        //    Response.ClearContent();
        //    Response.ClearHeaders();
        //    Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".pdf");
        //    Response.ContentType = "application/pdf";
        //    Response.Flush();
        //    Response.TransmitFile(Server.MapPath("~/Uploads/"+ SysPrefs.SubmissionFolder + "/" + fileName + ".pdf"));
        //    Response.End();
        //}

        public ActionResult ErrorMsg()
        {
            return View("~/Views/Shared/Error.cshtml");
        }


    }
}