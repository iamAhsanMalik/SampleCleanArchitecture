using DocumentFormat.OpenXml.Drawing;
using DHAAccounts.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace Accounting.Controllers
{
    public class AuthorizationController : Controller
    {
        ApplicationUser Profile = ApplicationUser.GetUserProfile();
        // GET: Authorization
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult CashTranx()
        {
            AuthorizationViewModel pModel = new AuthorizationViewModel();
            pModel.Type = "22BANKPMT";
            pModel.LabelType = "Cash Voucher Authorization";
            //pModel.table = BuildMainTable();
            return View("~/Views/Authorization/unpaid.cshtml", pModel);
        }

        public ActionResult TransferTranx()
        {
            AuthorizationViewModel pModel = new AuthorizationViewModel();
            pModel.Type = "22JOURNAL";
            pModel.LabelType = "Transfer Voucher Authorization";
            //pModel.table = BuildMainTable();
            return View("~/Views/Authorization/unpaid.cshtml", pModel);
        }

        public ActionResult DepositTranx()
        {
            AuthorizationViewModel pModel = new AuthorizationViewModel();
            pModel.Type = "22BANKDPT";
            pModel.LabelType = "Deposit Voucher Authorization";
            //pModel.table = BuildMainTable();
            return View("~/Views/Authorization/unpaid.cshtml", pModel);
        }

        public ActionResult HoldTranx()
        {
            AuthorizationViewModel pModel = new AuthorizationViewModel();
            pModel.Type = "22HOLDENTRY";
            pModel.LabelType = "Hold/Unpaid Voucher Authorization";
            //pModel.table = BuildMainTable();
            return View("~/Views/Authorization/unpaid.cshtml", pModel);
        }

        public ActionResult PaidTranx()
        {
            AuthorizationViewModel pModel = new AuthorizationViewModel();
            pModel.Type = "22PAIDENTRY";
            pModel.LabelType = "Paid Voucher Authorization";
            //pModel.table = BuildMainTable();
            return View("~/Views/Authorization/unpaid.cshtml", pModel);
        }

        public ActionResult CancelTranx()
        {
            AuthorizationViewModel pModel = new AuthorizationViewModel();
            pModel.Type = "22CANCELENTRY";
            pModel.LabelType = "Canceled Voucher Authorization";
            //pModel.table = BuildMainTable();
            return View("~/Views/Authorization/unpaid.cshtml", pModel);
        }
        [HttpPost]
        public ActionResult ShowAllVouchers(string ids)
        {
            AuthorizationViewModel pModel = new AuthorizationViewModel();
            int i = 1;
            string table = "";
            table += "<table border='1' width='100%' cellpadding='5'>";
            table += "<thead>";
            table += " <tr>";
            table += "<th style='background-color:#1c3a70; color:#FFF;'>Selected Vouchers</th>";
            table += "</tr>";
            table += "<thead>";
            table += "<tbody>";
            if (ids != "")
            {
                pModel.RefID = ids;
                string sqlList = "Select VoucherNumber, DatePosted, (select Description from ConfigItemsData where ConfigItemsData.ItemsDataCode=TypeiCode ) as TypeiCode from GLReferences Where isPosted=0 And GLReferenceID in(" + pModel.RefID + ")";
                DataTable transTable = DBUtils.GetDataTable(sqlList);
                if (transTable != null && transTable.Rows.Count > 0)
                {
                    foreach (DataRow RecordTable in transTable.Rows)
                    {
                        pModel.HtmlBody += i + ") &nbsp  " + "Voucher # " + RecordTable["VoucherNumber"].ToString() + " - " + RecordTable["TypeiCode"].ToString() + " - " + Common.toDateTime(RecordTable["DatePosted"]).ToShortDateString() + " is selected.<br>";
                        i++;
                    }
                    table += "<tr><td>" + pModel.HtmlBody + "</td></tr></tbody></table>";
                    table += "<div style='padding-top:15px:'><button type='submit' onclick='authorizeAllVouchers();' class='button-one right'>Authorise Selected Vouchers</button> <button type='submit' onclick='deleteAllVouchers();' class='button-one right'>Delete Selected Vouchers</button> <button type='submit' onclick='CancelVoucher();' class='button-one right'>Cancel</button></div>";
                    pModel.HtmlBody = table;
                    pModel.isPosted = true;
                }
            }

            return Json(pModel, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public ActionResult AuthorizeAllVouchers(string ids)
        {
            AuthorizationViewModel AModel = new AuthorizationViewModel();
            int i = 1;
            string table = "";
            string Typeicode = "";
            table += "<table border='1' width='100%' cellpadding='5'>";
            table += "<thead>";
            table += " <tr>";
            table += "<th style='background-color:#1c3a70; color:#FFF;'>Following vouchers has been authorized successfully.</th>";
            table += "</tr>";
            table += "<thead>";
            table += "<tbody>";
            if (ids != "")
            {
                AModel.RefID = ids;
                bool isAllowed = false;
                string sqlList = "Select VoucherNumber, DatePosted, TypeiCode ,(select Description from ConfigItemsData where ConfigItemsData.ItemsDataCode=TypeiCode ) as typeiCode from GLReferences Where isPosted=0 And GLReferenceID in(" + AModel.RefID + ")";
                DataTable transTable = DBUtils.GetDataTable(sqlList);
                if (transTable != null && transTable.Rows.Count > 0)
                {

                    foreach (DataRow RecordTable in transTable.Rows)
                    {
                        Typeicode = RecordTable["TypeiCode"].ToString();
                        AModel.HtmlBody += i + ") &nbsp  " + "Voucher # " + RecordTable["VoucherNumber"].ToString() + " - " + RecordTable["typeiCode"].ToString() + " - " + Common.toDateTime(RecordTable["DatePosted"]).ToShortDateString() + ".<br>";
                        i++;
                    }
                }
                SqlConnection con = Common.getConnection();
                string[] arr = ids.Split(',');
                foreach (string RefID in arr)
                {
                    string isPosted = DBUtils.executeSqlGetSingle("Select isPosted from GLReferences where [glreferenceid]= " + RefID + "");
                    if (isPosted == "False")
                    {
                        string PostedBy = DBUtils.executeSqlGetSingle("Select PostedBy from GLReferences where [glreferenceid]= " + RefID + "");
                        string updateGlReffrence = "UPDATE [dbo].[GLReferences] SET [isPosted] = @isPosted, [Authorizedby] = @Authorizedby, [Authorizeddate] = @Authorizeddate  where [glreferenceid] = @ReferenceNo";
                        if (Common.GetSysSettings("AutoTransferAuthorizationBSU") == "1")
                        {
                            isAllowed = true;
                        }
                        else
                        {
                            if (PostedBy != Profile.Id)
                            {
                                isAllowed = true;
                            }
                        }
                        if (isAllowed)
                        {
                            ////using ( SqlConnection con = Common.getConnection())
                            // //{
                            try
                            {
                                SqlCommand UpdateCommad = con.CreateCommand();
                                string myRefferenceNo = Common.toString(RefID);
                                if (myRefferenceNo != "")
                                {
                                    UpdateCommad.Parameters.AddWithValue("@isPosted", 1);
                                    UpdateCommad.Parameters.AddWithValue("@Authorizedby", Profile.Id);
                                    UpdateCommad.Parameters.AddWithValue("@Authorizeddate", DateTime.Now);
                                    UpdateCommad.Parameters.AddWithValue("@ReferenceNo", myRefferenceNo);
                                    UpdateCommad.CommandText = updateGlReffrence;
                                    UpdateCommad.ExecuteNonQuery();
                                    UpdateCommad.Parameters.Clear();
                                    if (Typeicode == "22CANCELENTRY")
                                    {
                                        DBUtils.ExecuteSQL("Update CustomerTransactions set Status = 'Cancelled' where CancelledGLReferenceID !=''");
                                    }
                                }
                                AModel.Message = Common.GetAlertMessage(0, "&nbsp &nbsp &nbsp Transfer voucher authorized successfully.");
                            }
                            catch
                            {

                            }

                            //// }
                        }

                        else
                        {
                            AModel.Message = Common.GetAlertMessage(1, "Sorry! same user can not authorize it.");
                        }
                    }

                }
                con.Close();
                table += "<tr><td>" + AModel.HtmlBody + "</td></tr></tbody></table>";
                AModel.HtmlBody = table;

            }
            else
            {
                AModel.Message = Common.GetAlertMessage(1, "Sorry! Vouchers not found for authorization.");
            }
            return Json(AModel, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public ActionResult DeleteAllVouchers(string ids)
        {
            AuthorizationViewModel pModel = new AuthorizationViewModel();
            int i = 1;
            string table = "";
            string Typeicode = "";
            table += "<table border='1' width='100%' cellpadding='5'>";
            table += "<thead>";
            table += " <tr>";
            table += "<th style='background-color:#1c3a70; color:#FFF;'>Following vouchers has been deleted successfully.</th>";
            table += "</tr>";
            table += "<thead>";
            table += "<tbody>";
            if (ids != "")
            {
                pModel.RefID = ids;
                bool isAllowed = false;
                string sqlList = "Select VoucherNumber, DatePosted, typeiCode ,(select Description from ConfigItemsData where ConfigItemsData.ItemsDataCode=TypeiCode ) as TypeiCode from GLReferences Where isPosted=0 And GLReferenceID in(" + pModel.RefID + ")";
                DataTable transTable = DBUtils.GetDataTable(sqlList);
                if (transTable != null && transTable.Rows.Count > 0)
                {

                    foreach (DataRow RecordTable in transTable.Rows)
                    {
                        Typeicode = RecordTable["typeiCode"].ToString();
                        pModel.HtmlBody += i + ") &nbsp  " + "Voucher # " + RecordTable["VoucherNumber"].ToString() + " - " + RecordTable["TypeiCode"].ToString() + " - " + Common.toDateTime(RecordTable["DatePosted"]).ToShortDateString() + ".<br>";
                        i++;
                    }
                }
                SqlConnection con = Common.getConnection();
                string[] arr = ids.Split(',');
                foreach (string RefID in arr)
                {
                    string refid = DBUtils.executeSqlGetSingle("Select GLReferenceID from GLReferences where [glreferenceid]= " + RefID + "");
                    if (Common.toString(refid) != "" && Common.toString(refid) != null)
                    {
                        string PostedBy = DBUtils.executeSqlGetSingle("Select PostedBy from GLReferences where [glreferenceid] = " + pModel.RefID + "");
                        if (Common.GetSysSettings("AutoTransferAuthorizationBSU") == "1")
                        {
                            isAllowed = true;
                        }
                        else
                        {
                            if (PostedBy != Profile.Id)
                            {
                                isAllowed = true;
                            }
                        }
                        if (isAllowed)
                        {
                            try
                            {
                                string SqlDeleteVoucher = "Delete From GLReferences where GLReferenceID=@GLReferenceID;";
                                string myRefferenceNo = Common.toString(RefID);
                                if (myRefferenceNo != "")
                                {
                                    //// using (SqlConnection con = Common.getConnection())
                                    /// / {
                                    SqlCommand Deletecommand = con.CreateCommand();
                                    string SqlDeletetransaction = "Delete from GLTransactions where GLReferenceID=@GLReferenceID";
                                    Deletecommand.Parameters.AddWithValue("@GLReferenceID", RefID);
                                    Deletecommand.CommandText = SqlDeletetransaction;
                                    Deletecommand.ExecuteNonQuery();
                                    Deletecommand.Parameters.Clear();
                                    Deletecommand.Parameters.AddWithValue("@GLReferenceID", myRefferenceNo);
                                    Deletecommand.CommandText = SqlDeleteVoucher;
                                    Deletecommand.ExecuteNonQuery();
                                    Deletecommand.Parameters.Clear();
                                    if (Typeicode == "22CANCELENTRY")
                                    {
                                        DBUtils.ExecuteSQL("Update CustomerTransactions set CancelledGLReferenceID = NULL where CancelledGLReferenceID = '" + RefID + "'");
                                    }
                                    //// }
                                }
                            }
                            catch
                            {

                            }
                        }
                        else
                        {
                            pModel.Message = Common.GetAlertMessage(1, "Sorry! same user can not delete it.");
                        }
                    }
                }
                con.Close();
                table += "<tr><td>" + pModel.HtmlBody + "</td></tr></tbody></table>";
                pModel.HtmlBody = table;
            }
            else
            {
                pModel.Message = Common.GetAlertMessage(1, "Sorry! Vouchers not found.");
            }
            return Json(pModel, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public ActionResult BuildMainTable(AuthorizationViewModel pModel)
        {
            string table = "";
            try
            {
                table += "<table id='tblMain' border='1' width='100%' cellpadding='5'>";
                table += "<thead>";
                table += " <tr>";
                table += "<th style='background-color:#007bff; color:#FFF;'><input id='chk' type='checkbox' value=" + 1 + " onclick='checkAll(this);'></th>";
                table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Voucher #</th>";
                table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Date</th>";
                table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Type</th>";
                table += "</tr>";
                table += "<thead>";
                table += "<tbody>";
                DataTable lstGLEntriesTemp = new DataTable();
                string sqlList = "Select GLReferenceID, VoucherNumber, ReferenceNo, DatePosted, PostedBy, (select Description from ConfigItemsData where ConfigItemsData.ItemsDataCode=TypeiCode ) as TypeiCode from GLReferences Where isPosted=0 ";
                if (pModel.Type != "")
                {
                    sqlList += " And TypeiCode= '" + pModel.Type + "'";
                }
                sqlList += " order by GLReferenceID DESC";
                lstGLEntriesTemp = DBUtils.GetDataTable(sqlList);
                if (lstGLEntriesTemp != null)
                {
                    if (lstGLEntriesTemp.Rows.Count > 0)
                    {
                        foreach (DataRow Item in lstGLEntriesTemp.Rows)
                        {
                            table += "<tr id='" + Common.toString(Item["GLReferenceID"]) + "'><td><input id='chk_" + Common.toString(Item["GLReferenceID"]) + "' type='checkbox' Class='case' value=" + Common.toString(Item["GLReferenceID"]) + " onclick='checkClick(this);'></td>";
                            table += "<td>" + Item["VoucherNumber"].ToString() + "</td>";
                            table += "<td>" + Item["DatePosted"].ToString().Substring(0, 10) + "</td>";
                            table += "<td>" + Item["TypeiCode"].ToString() + "</td></tr>";
                        }
                        pModel.isPosted = true;
                    }
                    else
                    {
                        pModel.isPosted = false;
                        pModel.Message = "No record found for authorization";
                    }
                }
                else
                {
                    pModel.isPosted = false;
                    pModel.Message = "No record found for authorization";
                }
                table += "</tbody>";
                table += "</table>";
                pModel.HtmlBody = table;
            }
            catch (Exception ex)
            {
                pModel.isPosted = false;
                pModel.Message = "An error occured: " + ex.ToString();
            }

            return Json(pModel, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult BuildVoucherTable(AuthorizationViewModel pModel)
        {
            if (!string.IsNullOrEmpty(Common.toString(pModel.GLReferenceID)))
            {
                //pGLTransViewModel.cpRefID = Security.DecryptQueryString(id);
                //pGLTransViewModel.HtmlTable1 = BuildDynamicTable1(Common.toInt(pGLTransViewModel.cpRefID));
                string Records = "Select VoucherNumber, AuthorizedBy, AuthorizedDate, DatePosted, PostedBy from GLReferences where isPosted=0 and GLReferenceID =" + pModel.GLReferenceID + "";
                DataTable transTable = DBUtils.GetDataTable(Records);
                if (transTable != null && transTable.Rows.Count > 0)
                {
                    foreach (DataRow RecordTable in transTable.Rows)
                    {
                        pModel.AuthorizedBy = Common.GetUserFullName(Common.toString(RecordTable["AuthorizedBy"]));
                        pModel.AuthorizedDate = Common.toDateTime(RecordTable["AuthorizedDate"]).ToShortDateString();
                        pModel.DatePosted = Common.toDateTime(RecordTable["DatePosted"]).ToShortDateString();
                        pModel.PostedBy = Common.GetUserFullName(Common.toString(RecordTable["PostedBy"]));
                        pModel.VoucherNumber = Common.toString(RecordTable["VoucherNumber"]);
                    }

                    string table = "";
                    table += "<div class='row'><div class='col-md-6'><div class='media'><div class='media-body'>";
                    table += "<h4 class='media-heading' style='color:darkred;'>" + @SysPrefs.SiteName + "</h4></div> </div> <br/><br/>";
                    table += "<div align ='center' class='text-center'><strong>Duplicate Voucher</strong></div><br/><br/>";
                    table += "<h4><b><u>Transfer Entry Voucher</u></b></h4></div><div class='col-md-6'>";
                    table += "<div style = 'border:#CCC 1px solid; padding:10px;'><table align='center'><tr>";
                    table += "<td><h5><b>Voucher No.</b></h5></td><td></td><td><h5><b>" + pModel.VoucherNumber + "</b></h5></td></tr><tr>";
                    table += "<td>Print Date:</td><td></td> <td>" + DateTime.Now.ToShortDateString() + "</td></tr><tr><td>Entry Date:</td><td></td>";
                    table += "<td>" + pModel.DatePosted + "</td></tr><tr><td>Authorize Date:</td><td></td><td>" + pModel.AuthorizedDate + "</td></tr><tr><td>Entry By:</td><td>";
                    table += "</td><td>" + pModel.PostedBy + "</td></tr><tr><td>Authorize By:</td><td></td><td>" + pModel.AuthorizedBy + "</td></tr></table></div></div></div>";
                    decimal cpTotalDebit = 0m;
                    decimal cpTotalCredit = 0m;
                    try
                    {
                        table += "<table border='1' width='100%' cellpadding='5'>";
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
                        if (pModel.GLReferenceID > 0)
                        {
                            DataTable objGLTransView = DBUtils.GetDataTable("select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate from GLTransactions where GLReferenceID='" + pModel.GLReferenceID.ToString() + "' order by LocalCurrencyAmount Desc");
                            if (objGLTransView != null && objGLTransView.Rows.Count > 0)
                            {
                                int iRowNumber = 0;
                                foreach (DataRow MyRow in objGLTransView.Rows)
                                {
                                    DataTable objGLReference = DBUtils.GetDataTable("select GLReferenceID, TypeiCode, ReferenceNo, Memo From GLReferences Where isPosted=0 and GLReferenceID='" + pModel.GLReferenceID.ToString() + "'");
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
                                        table += "<td valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";
                                        //string LC = "0.00";
                                        //if (!string.IsNullOrEmpty(MyRow["LocalCurrencyAmount"].ToString()))
                                        //{
                                        //    LC = MyRow["LocalCurrencyAmount"].ToString();
                                        //}
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
                                        //decimal LcAmount = Common.toDecimal(LC);
                                        //if (LcAmount < 0)
                                        //{
                                        //    table += "<td>&nbsp;</td>";
                                        //    table += "<td valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat((-1 * LcAmount)) + "</td>";
                                        //    cpTotalCredit = cpTotalCredit + (-1 * LcAmount);
                                        //}
                                        //else
                                        //{
                                        //    table += "<td valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                        //    table += "<td>&nbsp;</td>";
                                        //    cpTotalDebit = cpTotalDebit + LcAmount;
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
                                table += "<tfoot>";
                                table += "<tr>";
                                table += "<td colspan='2'></td>";
                                table += "<td style='background-color:#1c3a70; color:#FFF;'>Total:</td>";
                                table += "<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + DisplayUtils.GetSystemAmountFormat(cpTotalDebit) + "</td>";
                                table += "<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + DisplayUtils.GetSystemAmountFormat(cpTotalCredit) + "</td>";
                                table += "</tr>";
                                table += "</tfoot>";
                                table += "</table>";
                                table += "<br>";
                                table += "<h4> Attachments </h4>";
                                string getFileName = "Select FileName from GLReferencesAttachment where GLReferenceID='" + pModel.GLReferenceID + "'";
                                DataTable dtFileName = DBUtils.GetDataTable(getFileName);
                                if (dtFileName != null && dtFileName.Rows.Count > 0)
                                {
                                    foreach (DataRow drFileName in dtFileName.Rows)
                                    {
                                        // pGLTransViewModel.Attachments += Common.toString(drFileName["FileName"]) + ", ";
                                        table += "<a href ='/Uploads/" + SysPrefs.SubmissionFolder + "/" + drFileName["FileName"].ToString() + "', new { target = '_blank' }>" + Common.toString(drFileName["FileName"]) + "</a>" + "<br>";
                                    }
                                }
                                table += "<br><br>";
                                table += "<div style='padding-top:15px:'><button type='submit' value=" + pModel.GLReferenceID.ToString() + " onclick='authorizeVoucher(this);' class='button-one right'>Authorise Voucher</button> <button type='submit' value=" + pModel.GLReferenceID.ToString() + " onclick='deleteVoucher(this);' class='button-one right'>Delete</button> <button type='submit' onclick='CancelVoucher();' class='button-one right'>Cancel</button></div>";
                                pModel.HtmlBody = table;
                                pModel.isPosted = true;
                            }
                            else
                            {
                                pModel.isPosted = false;
                                pModel.Message = "Invalid info. Provide valid voucher number.";
                            }
                        }
                        else
                        {
                            pModel.isPosted = false;
                            pModel.Message = "Invalid info. Provide valid voucher number.";
                        }

                    }
                    catch (Exception ex)
                    {
                        pModel.isPosted = false;
                        pModel.Message = "An error occured: " + ex.ToString();
                    }
                }
                else
                {
                    pModel.Message = "Voucher already authorized or deleted.";
                }
            }

            return Json(pModel, JsonRequestBehavior.AllowGet);
        }
        public ActionResult AuthorizeVoucher(AuthorizationViewModel AModel)
        {
            bool isAllowed = false;
            string PostedBy = DBUtils.executeSqlGetSingle("Select PostedBy from GLReferences where [glreferenceid]=" + AModel.GLReferenceID + "");
            string updateGlReffrence = "UPDATE [dbo].[GLReferences] SET [isPosted] = @isPosted, [Authorizedby] = @Authorizedby, [Authorizeddate] = @Authorizeddate  where [glreferenceid]=@ReferenceNo";
            if (Common.GetSysSettings("AutoTransferAuthorizationBSU") == "1")
            {
                isAllowed = true;
            }
            else
            {
                //if (PostedBy != Profile.Id)
                //{
                //    isAllowed = true;
                //}
                isAllowed = true;
            }
            if (isAllowed)
            {
                SqlConnection con = Common.getConnection();
                ////using (SqlConnection con = Common.getConnection())
                ////{
                try
                {
                    SqlCommand UpdateCommad = con.CreateCommand();
                    string myRefferenceNo = Common.toString(AModel.GLReferenceID);
                    if (myRefferenceNo != "")
                    {
                        UpdateCommad.Parameters.AddWithValue("@isPosted", 1);
                        UpdateCommad.Parameters.AddWithValue("@Authorizedby", Profile.Id);
                        UpdateCommad.Parameters.AddWithValue("@Authorizeddate", DateTime.Now);
                        UpdateCommad.Parameters.AddWithValue("@ReferenceNo", myRefferenceNo);
                        UpdateCommad.CommandText = updateGlReffrence;
                        UpdateCommad.ExecuteNonQuery();
                        UpdateCommad.Parameters.Clear();
                        if (AModel.Type == "22CANCELENTRY")
                        {
                            DBUtils.ExecuteSQL("Update CustomerTransactions set Status = 'Cancelled' where CancelledGLReferenceID = '" + AModel.GLReferenceID + "'");
                        }
                    }
                    AModel.Message = Common.GetAlertMessage(0, "&nbsp &nbsp &nbsp Transfer voucher authorized successfully.");
                }
                catch
                {

                }

                ////  }
                con.Close();
            }
            else
            {
                AModel.Message = Common.GetAlertMessage(1, "Sorry! same user can not authorize it.");
            }
            return Json(AModel, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DeleteVoucher(AuthorizationViewModel pModel)
        {
            bool isAllowed = false;
            string PostedBy = DBUtils.executeSqlGetSingle("Select PostedBy from GLReferences where [glreferenceid]=" + pModel.GLReferenceID + "");
            if (Common.GetSysSettings("AutoTransferAuthorizationBSU") == "1")
            {
                isAllowed = true;
            }
            else
            {
                if (PostedBy != Profile.Id)
                {
                    isAllowed = true;
                }
            }
            if (isAllowed)
            {
                try
                {
                    SqlConnection con = Common.getConnection();
                    string SqlDeleteVoucher = "Delete From GLReferences where GLReferenceID=@GLReferenceID;";
                    string myRefferenceNo = Common.toString(pModel.GLReferenceID);
                    if (myRefferenceNo != "")
                    {
                        ////using (SqlConnection con = Common.getConnection())
                        // //{
                        SqlCommand Deletecommand = con.CreateCommand();
                        string SqlDeletetransaction = "Delete from GLTransactions where GLReferenceID=@GLReferenceID";
                        Deletecommand.Parameters.AddWithValue("@GLReferenceID", pModel.GLReferenceID);
                        Deletecommand.CommandText = SqlDeletetransaction;
                        Deletecommand.ExecuteNonQuery();
                        Deletecommand.Parameters.Clear();
                        Deletecommand.Parameters.AddWithValue("@GLReferenceID", myRefferenceNo);
                        Deletecommand.CommandText = SqlDeleteVoucher;
                        Deletecommand.ExecuteNonQuery();
                        Deletecommand.Parameters.Clear();
                        if (pModel.Type == "22CANCELENTRY")
                        {
                            string status = "";
                            //  status = PostingUtils.GetPreviousTranxStatus(Common.toString(pModel.GLReferenceID));
                            //   DBUtils.ExecuteSQL("Update CustomerTransactions set Status='"+ status + "', CancelledGLReferenceID = NULL where CancelledGLReferenceID = '" + pModel.GLReferenceID + "'");
                            DBUtils.ExecuteSQL("Update CustomerTransactions set CancelledGLReferenceID = NULL where CancelledGLReferenceID = '" + pModel.GLReferenceID + "'");
                        }
                        con.Close();
                        ////  }
                    }
                    pModel.Message = Common.GetAlertMessage(0, "&nbsp &nbsp &nbsp Voucher deleted succcessfully.");
                }
                catch (Exception ex)
                {
                    pModel.Message = Common.GetAlertMessage(1, "Error occured.");
                }
            }
            else
            {
                pModel.Message = Common.GetAlertMessage(1, "Sorry! same user can not deleted it.");
            }
            return Json(pModel, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateVoucher(AuthorizationViewModel pModel)
        {
            return RedirectToAction("GeneralJournal", "GLGeneralJournal");
        }




    }
}