using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Spreadsheet;
using DHAAccounts.Models;


namespace Accounting.Controllers
{
    public class DashBoardController : Controller
    {
        ApplicationUser profile = ApplicationUser.GetUserProfile();
        // GET: DashBoard
        public ActionResult Index()
        {
            //string sql = "select GLTransactions.AddedDate as Date,GLTransactions.GLReferenceID as ID FROM GLTransactions, GLReferences where  GLTransactions.glreferenceid = GLReferences.glreferenceid  and TypeiCode = '22CANCELENTRY' and  GLReferences.isPosted = 1";
            //DataTable dt = DBUtils.GetDataTable(sql);
            //if (dt!=null && dt.Rows.Count>0)
            //{
            //    foreach (DataRow dr in dt.Rows)
            //    {
            //        string sqlTrans = "update CustomerTransactions set CancelledDate='" + Common.toString(dr["Date"]) + "' where CancelledGLReferenceID='" + Common.toString(dr["ID"]) + "'";
            //        DBUtils.ExecuteSQL(sqlTrans);
            //    }

            //}

            //if (!System.IO.Directory.Exists("/Uploads/" + SysPrefs.SubmissionFolder))
            //{
            //    System.IO.Directory.CreateDirectory("/Uploads/" + SysPrefs.SubmissionFolder);
            //}
            //if (!System.IO.Directory.Exists("/Uploads/" + SysPrefs.SubmissionFolder + "/profiles"))
            //{
            //    System.IO.Directory.CreateDirectory("/Uploads/" + SysPrefs.SubmissionFolder + "/profiles");
            //}
            //if (!System.IO.Directory.Exists("/Uploads/" + SysPrefs.SubmissionFolder + "/Vouchers"))
            //{
            //    System.IO.Directory.CreateDirectory("/Uploads/" + SysPrefs.SubmissionFolder + "/Vouchers");
            //}
            //if (!System.IO.Directory.Exists("/Uploads/" + SysPrefs.SubmissionFolder+"/DownTemp"))
            //{
            //    System.IO.Directory.CreateDirectory("/Uploads/" + SysPrefs.SubmissionFolder + "/DownTemp");
            //}
            //if (!System.IO.Directory.Exists("/Uploads/" + SysPrefs.SubmissionFolder + "/UpTemp"))
            //{
            //    System.IO.Directory.CreateDirectory("/Uploads/" + SysPrefs.SubmissionFolder + "/UpTemp");
            //}

            return View("~/Views/DashBoard/Index.cshtml");
        }

        public ActionResult RecentTransactions()
        {
            string table = "";
            string Memo = "";
            table += "<table id='tblMain' border='1' width='172%' cellpadding='5'>";
            table += "<thead>";
            table += "<tr>";
            table += "<th style='background-color:#007bff; color:#FFF; width:3%;'>Date</th>";
            table += "<th style='background-color:#007bff; color:#FFF; width:8%;'>V.No</th>";
            table += "<th style='background-color:#007bff; width:35%; color:#FFF;'>Description</th>";
            table += "<th style='background-color:#007bff; width:20%; color:#FFF;'>Account Title</th>";
            table += "<th style='background-color:#007bff; color:#FFF; width:3%;'>Curr</th>";
            table += "<th style='background-color:#007bff; color:#FFF; width:3%;'>F/C</th>";
            table += "<th style='background-color:#007bff; color:#FFF; width:3%;'>Debit</th>";
            table += "<th style='background-color:#007bff; color:#FFF; width:3%;'>Credit</th>";
            table += "</tr>";
            table += "</thead>";
            table += "<tbody>";
            string mySql = "Select top(10) GLReferenceID, DatePosted, VoucherNumber, Memo,(select Description from ConfigItemsData where GLReferences.TypeiCode= ConfigItemsData.ItemsDataCode) as TypeiCode from GLReferences order by GLReferenceID Desc";
            DataTable dtReferences = DBUtils.GetDataTable(mySql);
            if (dtReferences != null)
            {
                foreach (DataRow MyRow in dtReferences.Rows)
                {
                    table += "<tr>";
                    table += "<td valign='top'>" + Common.toDateTime(MyRow["DatePosted"]).ToShortDateString() + "</td>";
                    table += "<td valign='top' style='font-size:xx-small'>" + "<a href=\"javascript: showDetails('/Inquiries/GLTransView?Id=" + Security.EncryptQueryString(MyRow["GLReferenceID"].ToString()) + "','Voucher View')\")>" + MyRow["VoucherNumber"].ToString() + "</a><br>" + "(" + MyRow["TypeiCode"].ToString() + ")" + "</td>";
                    if (MyRow["Memo"].ToString().ToCharArray().Count() > 100)
                    {
                        Memo = MyRow["Memo"].ToString().Substring(0, 100);
                    }
                    else
                    {
                        Memo = MyRow["Memo"].ToString();
                    }
                    table += "<td valign='top'>" + Memo + "</td>";
                    string sqlTranx = "Select GLTransactions.GLReferenceID, (Select AccountName From GLChartOfAccounts Where GLTransactions.AccountCode = GLChartOfAccounts.AccountCode) as AccountTitle, GLTransactions.ForeignCurrencyAmount, GLTransactions.LocalCurrencyAmount, GLTransactions.ForeignCurrencyISOCode FROM GLTransactions Where TransactionID > 0 and GLReferenceID='" + MyRow["GLReferenceID"].ToString() + "' order by LocalCurrencyAmount Desc";
                    DataTable dtTranx = DBUtils.GetDataTable(sqlTranx);
                    if (dtTranx != null)
                    {
                        string accountTitle = "";
                        string currency = "";
                        string Fc = "";
                        string debitamt = "";
                        string creditamt = "";
                        foreach (DataRow myRow in dtTranx.Rows)
                        {
                            accountTitle += myRow["AccountTitle"].ToString() + "<br>";
                            currency += myRow["ForeignCurrencyISOCode"].ToString() + "<br>";
                            decimal FcAmt = Math.Abs(Math.Round(Convert.ToDecimal(myRow["ForeignCurrencyAmount"].ToString()), 2));
                            if (FcAmt != 0.00m)
                            {
                                Fc += FcAmt + "<br>";
                            }
                            string LC = "0.00";
                            decimal LcAmount = 0;
                            if (!string.IsNullOrEmpty(myRow["LocalCurrencyAmount"].ToString()))
                            {
                                LC = myRow["LocalCurrencyAmount"].ToString();

                            }
                            LcAmount = DisplayUtils.getRoundDigit(LC);
                            if (LcAmount < 0)
                            {
                                debitamt += "<br>";
                                creditamt += DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "<br>";
                            }
                            else
                            {

                                debitamt += DisplayUtils.GetSystemAmountFormat(LcAmount) + "<br>";
                                creditamt += "<br>";
                            }
                        }
                        table += "<td valign='top' style='font-size:xx-small'>" + accountTitle + "</td>";
                        table += "<td valign='top' align='center' style='font-size:xx-small'>" + currency + "</td>";
                        table += "<td valign='top' align='right' style='font-size:xx-small'>" + Fc + "</td>";
                        table += "<td valign='top' align='right' style='font-size:xx-small'>" + debitamt + "</td>";
                        table += "<td valign='top' align='right' style='font-size:xx-small'>" + creditamt + "</td>";
                    }
                    table += "</tr>";
                }
                table += "</tbody>";
                table += "</table>";
            }
            return Json(table, JsonRequestBehavior.AllowGet);
        }
        public ActionResult GetRates()
        {
            string table = "";
            table += "<table id='tblMain' border='1' width='100%' cellpadding='5'>";
            table += "<thead>";
            table += "<tr>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Code</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Currency Name</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Exchange Rate</th>";
            table += "</tr>";
            table += "</thead>";
            table += "<tbody>";
            string mySql = "Select CurrencyISOCode ,(Select CurrencyName from ListCurrencies where ExchangeRates.CurrencyISOCode=ListCurrencies.CurrencyISOCode and CurrencyStatus = '1') as CurrencyName, ExchangeRate from ExchangeRates where AddedDate='" + Common.toDateTime(SysPrefs.PostingDate).ToString("MM/dd/yyyy 00:00:00") + "' Order By CurrencyISOCode";
            DataTable dtRates = DBUtils.GetDataTable(mySql);
            if (dtRates != null)
            {
                foreach (DataRow MyRow in dtRates.Rows)
                {
                    table += "<tr>";
                    table += "<td align='left'>" + MyRow["CurrencyISOCode"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["CurrencyName"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["ExchangeRate"].ToString() + "</td>";
                    table += "</tr>";

                }
                table += "</tbody>";
                table += "</table>";
            }
            return Json(table, JsonRequestBehavior.AllowGet);
        }
        public ActionResult RecentImportedTransactions()
        {
            string table = "";
            table += "<table id='tblMain' border='1' width='100%' cellpadding='5'>";
            table += "<thead>";
            table += "<tr>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Date</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Inv.#</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Payment #</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Agent</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Sender</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Bene</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Curr</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Sending</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Receiving</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Status</th>";
            table += "</tr>";
            table += "</thead>";
            table += "<tbody>";
            //string mySql = "Select CustomerTransactionID, PostingDate, PaymentNo, AgentPrefix, SenderName,Recipient,Phone,Status,Currency,FC_Amount,TransactionID,CustomerID,RecevingCountry,(Select top(1) GLReferenceID from GLTransactions where CustomerTransactions.PaymentNo=GLTransactions.TransactionNumber) as GLReferenceID,Pounds FROM CustomerTransactions order by CustomerTransactionID desc";
            string mySql = "Select top 10 CustomerTransactionID, PostingDate, PaymentNo, AgentPrefix, SenderName,Recipient,Phone,Status,Currency,FC_Amount,TransactionID,CustomerID,RecevingCountry,Pounds, HoldGLReferenceID, PaidGLReferenceID, CancelledGLReferenceID FROM CustomerTransactions order by CustomerTransactionID desc";
            DataTable dtRates = DBUtils.GetDataTable(mySql);
            if (dtRates != null)
            {
                foreach (DataRow MyRow in dtRates.Rows)
                {
                    table += "<tr>";
                    table += "<td align='left'>" + Common.toDateTime(MyRow["PostingDate"]).ToShortDateString() + "</td>";
                    string GLReferenceID = "";
                    if (Common.toString(MyRow["CancelledGLReferenceID"]).Trim() != "")
                    {
                        GLReferenceID = Common.toString(MyRow["CancelledGLReferenceID"]).Trim();
                    }
                    else if (Common.toString(MyRow["PaidGLReferenceID"]).Trim() != "")
                    {
                        GLReferenceID = Common.toString(MyRow["PaidGLReferenceID"]).Trim();
                    }
                    else
                    {
                        GLReferenceID = Common.toString(MyRow["HoldGLReferenceID"]).Trim();
                    }

                    table += "<td align='left'>" + "<a href=\"javascript: showDetails('/Inquiries/GLTransView?Id=" + Security.EncryptQueryString(GLReferenceID) + "','Voucher View')\")>" + MyRow["TransactionID"].ToString() + "</a></td>";
                    table += "<td align='left'>" + MyRow["PaymentNo"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["AgentPrefix"].ToString() + "</td>";
                    table += "<td align='left' style='font-size:smaller'>" + MyRow["SenderName"].ToString() + "</td>";
                    table += "<td align='left' style='font-size:smaller'>" + MyRow["Recipient"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["Currency"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["FC_Amount"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["Pounds"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["Status"].ToString() + "</td>";
                    table += "</tr>";

                }
                table += "</tbody>";
                table += "</table>";
            }
            return Json(table, JsonRequestBehavior.AllowGet);
        }
        public ActionResult PendingAuthorizations()
        {
            string table = "";
            table += "<table id='tblMain' border='1' width='100%' cellpadding='5'>";
            table += "<thead>";
            table += "<tr>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Date</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Type</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Voucher #</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Invoice #</th>";
            table += "<th style='background-color:#007bff; color:#FFF;'>Payment #</th>";
            table += "</tr>";
            table += "</thead>";
            table += "<tbody>";
            string mySql = "Select DatePosted ,(select Description from ConfigItemsData where GLReferences.TypeiCode= ConfigItemsData.ItemsDataCode) as TypeiCode, VoucherNumber,(Select TransactionID from CustomerTransactions where GLReferences.GLReferenceID = CustomerTransactions.HoldGLReferenceID) as InvoiceNo,(Select PaymentNo from CustomerTransactions where GLReferences.GLReferenceID=CustomerTransactions.HoldGLReferenceID) as PaymentNo from GLReferences where isPosted = '0'";
            DataTable dtRates = DBUtils.GetDataTable(mySql);
            if (dtRates != null)
            {
                foreach (DataRow MyRow in dtRates.Rows)
                {
                    table += "<tr>";
                    table += "<td align='left'>" + Common.toDateTime(MyRow["DatePosted"]).ToShortDateString() + "</td>";
                    table += "<td align='left'>" + MyRow["TypeiCode"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["VoucherNumber"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["InvoiceNo"].ToString() + "</td>";
                    table += "<td align='left'>" + MyRow["PaymentNo"].ToString() + "</td>";
                    table += "</tr>";

                }
                table += "</tbody>";
                table += "</table>";
            }
            return Json(table, JsonRequestBehavior.AllowGet);
        }
        public ActionResult ReadSql()
        {
            try
            {
                string Html = "<table cellspacing='0' cellpadding='2' style='border-collapse: collapse;border: 1px solid #ccc;font-size: 9pt;'>";
                Html += "<tr>";
                Html += "<th style='background-color: #B8DBFD;border: 1px solid #ccc'>Customer Id</th>";
                Html += "<th style='background-color: #B8DBFD;border: 1px solid #ccc'>Name</th>";
                Html += "<th style='background-color: #B8DBFD;border: 1px solid #ccc'>Country</th>";
                Html += "</tr>";
                Html += "<tr>";
                Html += "<td style='width:120px;border: 1px solid #ccc'>1</td>";
                Html += "<td style='width:150px;border: 1px solid #ccc'>John Hammond</td>";
                Html += "<td style='width:120px;border: 1px solid #ccc'>United States</td>";
                Html += "</tr>";
                Html += "</table>";
                Response.Clear();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment;filename=HTML.xls");
                Response.Charset = "";
                Response.ContentType = "application/vnd.ms-excel";
                Response.Output.Write(Html);
                Response.Flush();
                Response.End();













                //DBUtils.ExecuteSQL("update CustomerTransactions set PostingDate = '2020-08-20 14:34:00' where CustomerTransactionID = 104801");
                // DBUtils.ExecuteSQL("update CustomerTransactions set PostingDate = '2020-08-14 00:00:00' where CustomerTransactionID = 137542");

                // string sql = "select GLTransactions.AddedDate as Date,GLTransactions.GLReferenceID as ID FROM GLTransactions, GLReferences where  GLTransactions.glreferenceid = GLReferences.glreferenceid  and TypeiCode = '22CANCELENTRY' and  GLReferences.isPosted = 1";
                //DataTable dt = DBUtils.GetDataTable(sql);
                //if (dt != null && dt.Rows.Count > 0)
                //{
                //    foreach (DataRow dr in dt.Rows)
                //    {
                //        if (Common.toString(dr["Date"])!="") {
                //            string sqlTrans = "update CustomerTransactions set CancelledDate='" + Common.toDateTime(Common.toString(dr["Date"])).ToString("yyyy-MM-dd 00:00:00") + "' where CancelledGLReferenceID='" + Common.toString(dr["ID"]) + "'";
                //            DBUtils.ExecuteSQL(sqlTrans);
                //            Response.Write("-----Successfully execute      " + sqlTrans + "          -----" + "<br/>");
                //        }

                //    }

                //}



                //string ExcelFilePath = Server.MapPath("~/Uploads/ReadSql.txt");
                //string[] lines = System.IO.File.ReadAllLines(ExcelFilePath);
                //foreach (string line in lines)
                //{
                //    Response.Write(line+"<br/>");
                //    bool OK=DBUtils.ExecuteSQL(line);
                //    if (OK)
                //    {
                //        Response.Write("-----Successfully execute" +"-----"+ "<br/>");
                //    }
                //}
                Response.End();
            }
            catch (Exception ex)
            {
            }
            return null;
        }
    }
}