using DHAAccounts.Models;
using Kendo.Mvc.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace Accounting.Controllers
{
    public class GLGeneralJournalController : Controller
    {
        //  private iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
        ApplicationUser profile = ApplicationUser.GetUserProfile();
        private AppDbContext _dbcontext = new AppDbContext();
        ////private AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());
        const string TypeiCode = "22JOURNAL";
        // GET: GernalJournal
        public ActionResult Index()
        {
            return View();
        }
        // GET: GeneralJournal
        #region GeneralJournal Edit
        public ActionResult GeneralJournalEdit()
        {
            PostingViewModel postingViewModel = new PostingViewModel();
            postingViewModel.isPosted = true;
            return View("~/Views/Postings/GeneralJournalEdit.cshtml", postingViewModel);
        }
        [HttpPost]
        public ActionResult GeneralJournalEdit(string VoucherNumber)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            PostingViewModel postingViewModel = new PostingViewModel();
            if (!string.IsNullOrEmpty(VoucherNumber))
            {
                string mySql = "Select GLReferenceID,DatePosted from GLReferences";
                if (!string.IsNullOrEmpty(VoucherNumber))
                {
                    mySql += " where VoucherNumber='" + VoucherNumber + "' and TypeiCode in ('22JOURNAL', '22BANKPMT')";
                }
                int id = 0;
                DataTable dtVoucher = DBUtils.GetDataTable(mySql);
                if (dtVoucher != null && dtVoucher.Rows.Count > 0)
                {
                    id = Common.toInt(dtVoucher.Rows[0]["GLReferenceID"]);
                    postingViewModel.DatePosted = Common.toDateTime(dtVoucher.Rows[0]["DatePosted"]);
                }
                //int id = Common.toInt(DBUtils.executeSqlGetSingle(mySql));
                if (id > 0)
                {
                    //GLReferences gLReference = db.GLReferences.Find(Common.toInt(id));
                    //if (gLReference != null)
                    //{
                    //    FillData(Common.toInt(id));
                    //    postingViewModel.GLReferenceID = Common.toInt(id);
                    //    postingViewModel.cpUserRefNo = Common.toString(gLReference.GLReferenceID);
                    //    postingViewModel.cpActionType = "E";
                    //    postingViewModel.cpRefID = Common.toString(gLReference.GLReferenceID);
                    //    postingViewModel.table = BuildDynamicTable(postingViewModel.cpUserRefNo);
                    //    postingViewModel.isPosted = false;
                    //    bool sqlUpdateEntry = DBUtils.ExecuteSQL("Update GLReferencesAttachment set IsDeleted = 0");
                    //    string getFileName = "Select * from GLReferencesAttachment where GLReferenceID='" + postingViewModel.cpUserRefNo + "' and IsDeleted = 0";
                    //    DataTable dtFileName = DBUtils.GetDataTable(getFileName);
                    //    if (dtFileName != null && dtFileName.Rows.Count > 0)
                    //    {
                    //        foreach (DataRow drFileName in dtFileName.Rows)
                    //        {
                    //            // pGLTransViewModel.Attachments += Common.toString(drFileName["FileName"]) + ", ";
                    //            postingViewModel.Attachments += "<a href ='/Uploads/Vouchers/" + drFileName["FileName"].ToString() + "', new { target = '_blank' }>" + Common.toString(drFileName["FileName"]) + "</a> <a href='javascript: deleteAttachment(" + Common.toString(drFileName["AttachmentID"]) + ")'><i class='fa fa-remove'></i></a> " + " ,";
                    //        }
                    //    }
                    //}
                    string sql = "select * from GLReferences where GLReferenceID=" + id;
                    DataTable dt = DBUtils.GetDataTable(sql);
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        FillData(Common.toInt(id));
                        postingViewModel.GLReferenceID = Common.toInt(id);
                        postingViewModel.cpUserRefNo = Common.toString(dt.Rows[0]["GLReferenceID"]);
                        postingViewModel.cpActionType = "E";
                        postingViewModel.cpRefID = Common.toString(dt.Rows[0]["GLReferenceID"]);
                        postingViewModel.table = BuildDynamicTable(postingViewModel.cpUserRefNo);
                        postingViewModel.isPosted = false;
                        bool sqlUpdateEntry = DBUtils.ExecuteSQL("Update GLReferencesAttachment set IsDeleted = 0");
                        string getFileName = "Select * from GLReferencesAttachment where GLReferenceID='" + postingViewModel.cpUserRefNo + "' and IsDeleted = 0";
                        DataTable dtFileName = DBUtils.GetDataTable(getFileName);
                        if (dtFileName != null && dtFileName.Rows.Count > 0)
                        {
                            foreach (DataRow drFileName in dtFileName.Rows)
                            {
                                // pGLTransViewModel.Attachments += Common.toString(drFileName["FileName"]) + ", ";
                                postingViewModel.Attachments += "<a href ='/Uploads/" + SysPrefs.SubmissionFolder + "/Vouchers/" + drFileName["FileName"].ToString() + "', new { target = '_blank' }>" + Common.toString(drFileName["FileName"]) + "</a> <a href='javascript: deleteAttachment(" + Common.toString(drFileName["AttachmentID"]) + ")'><i class='fa fa-remove'></i></a> " + " ,";
                            }
                        }
                    }
                    else
                    {
                        postingViewModel.isPosted = true;
                        postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Invalid Id provided");
                    }

                }
                else
                {
                    postingViewModel.isPosted = true;
                    postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Invalid Id provided");
                }
            }
            else
            {
                postingViewModel.isPosted = true;
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Please provide voucher no.");
            }
            return View("~/Views/Postings/GeneralJournalEdit.cshtml", postingViewModel);
        }
        #endregion

        public ActionResult GeneralJournal(int? id)
        {
            PostingViewModel postingViewModel = new PostingViewModel();
            Guid guid = Guid.NewGuid();
            postingViewModel.cpUserRefNo = guid.ToString();
            Session["cpUserRefNo"] = postingViewModel.cpUserRefNo;
            //if (id > 0)
            //{
            //    //if (cpActionType == "E")
            //    //{
            //    //    FillData();
            //    //}
            //    FillData(Common.toInt(id));
            //    postingViewModel.GLReferenceID = Common.toInt(id);
            //    GLReferences gLReference = db.GLReferences.Find(Common.toInt(id));
            //    postingViewModel.cpUserRefNo = gLReference.ReferenceNo;
            //}
            postingViewModel.table = BuildDynamicTable(postingViewModel.cpUserRefNo);
            return View("~/Views/Postings/GeneralJournal.cshtml", postingViewModel);
        }
        //Get GeneralJournalEdit
        public ActionResult ShowMessage(string cpUserRefNo)
        {
            if (cpUserRefNo != "")
            {
                cpUserRefNo = Security.DecryptQueryString(cpUserRefNo);
                cpUserRefNo = cpUserRefNo.Replace("\0", "");
            }
            PostingViewModel postingViewModel = PostingUtils.ShowMessage(cpUserRefNo);
            return View("~/Views/Postings/Message.cshtml", postingViewModel);
        }
        public JsonResult GetDebitCashAjax([DataSourceRequest] DataSourceRequest request)
        {
            var items = PostingUtils.GetDebitCashAjax();
            return Json(items, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetBankAccounts([DataSourceRequest] DataSourceRequest request)
        {
            var bankAccounts = PostingUtils.GetBankAccounts();
            return Json(bankAccounts, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetGeneralAccounts([DataSourceRequest] DataSourceRequest request)
        {
            var bankAccount = PostingUtils.GetGeneralAccounts();
            return Json(bankAccount, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetAccountBalance(string pAccountCode)
        {
            string LcAmounts = PostingUtils.GetAccountBalance(pAccountCode);
            return Json(LcAmounts);
        }

        public JsonResult GetAccountDetails(string pAccountCode)
        {
            AccountDetails objAccount = new AccountDetails();

            objAccount = PostingUtils.GetAccountDetails(pAccountCode);
            return Json(objAccount, JsonRequestBehavior.AllowGet);
        }
        protected static string BuildDynamicTable(string pUserRefNo)
        {
            PostingViewModel pPostingModel = new PostingViewModel();
            string table = PostingUtils.BuildDynamicTable(pUserRefNo);
            return table;
        }
        [HttpPost]
        //public ActionResult AddUpdate(string cashBal, string GeneralAccountCode, string BalLC, string BalFC, string pMemo, string pDrCr, string pFC, string pFCRate, string pLC, string cpRefNo, string fcISO, string lcISO)
        public ActionResult AddUpdate(AddUpdateViewModel viewModel)
        {
            ReturnViewModel returnModel = new ReturnViewModel();

            returnModel = PostingUtils.AddUpdate(viewModel, TypeiCode);
            return Json(returnModel, JsonRequestBehavior.AllowGet);
        }
        private static ReturnViewModel FillData(int pGLReferenceID)
        {
            ReturnViewModel pReturnViewModel = new ReturnViewModel();
            if (pGLReferenceID != null && pGLReferenceID > 0)
            {
                pReturnViewModel = PostingUtils.FillData(pGLReferenceID);
            }
            else
            {
                pReturnViewModel.ErrMessage = "Invalid GL Reference ID. please provide Reference ID";
                pReturnViewModel.isPosted = false;
            }

            return pReturnViewModel;
        }
        [HttpPost]
        public ActionResult ProceedJournalEntry(PostingViewModel pPostingModel, IEnumerable<HttpPostedFileBase> files)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            pPostingModel.isPosted = false;
            pPostingModel = PostingUtils.ProceedJournalEntry(pPostingModel);
            if (pPostingModel.isPosted)
            {
                if (files != null)
                {
                    SqlCommand InsertCommand = null;
                    SqlConnection con = Common.getConnection();

                    foreach (var file in files)
                    {
                        var fileName = Path.GetFileName(file.FileName);
                        var physicalPath = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/"), fileName);
                        file.SaveAs(physicalPath);
                        string insertAttachment = "Insert into GLReferencesAttachment (GLReferenceID,FileName,AddedBy,AddedDate) Values (@GLReferenceID,@FileName,@AddedBy,@AddedDate)";
                        ////SqlCommand InsertCommand = null;
                        ////using (SqlConnection con = Common.getConnection())
                        ////{
                        try
                        {
                            InsertCommand = con.CreateCommand();
                            InsertCommand.Parameters.AddWithValue("@GLReferenceID", pPostingModel.GLReferenceID);
                            InsertCommand.Parameters.AddWithValue("@FileName", fileName);
                            InsertCommand.Parameters.AddWithValue("@AddedBy", profile.Id);
                            InsertCommand.Parameters.AddWithValue("@AddedDate", Common.toDateTime(SysPrefs.PostingDate));
                            InsertCommand.CommandText = insertAttachment;
                            InsertCommand.ExecuteNonQuery();
                            InsertCommand.Parameters.Clear();
                        }
                        catch (Exception ex)
                        {
                            pPostingModel.ErrMessage = Common.GetAlertMessage(0, "Error Occured" + ex.ToString());
                        }
                        //// }
                    }
                    con.Close();
                }
                pPostingModel.ErrMessage = Common.GetAlertMessage(0, pPostingModel.ErrMessage);
                return RedirectToAction("ShowMessage", "GLGeneralJournal", new { cpUserRefNo = Security.EncryptQueryString(Common.toString(pPostingModel.GLReferenceID)) });
            }
            else
            {
                pPostingModel.ErrMessage = Common.GetAlertMessage(1, pPostingModel.ErrMessage);
            }
            return View("~/Views/Postings/GeneralJournal.cshtml", pPostingModel);
        }

        [HttpPost]
        public ActionResult ProceedJournalEntryEdit(PostingViewModel pPostingModel, IEnumerable<HttpPostedFileBase> files)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            ApplicationUser Profile = ApplicationUser.GetUserProfile();
            if (Common.ValidatePinCode(Profile.Id, pPostingModel.PinCode))
            {
                if (Common.ValidateFiscalYearDate(Common.toDateTime(pPostingModel.DatePosted)))
                {
                    string sqlupdate = "Update GLEntriesTemp set AddedDate = '" + Common.toDateTime(pPostingModel.DatePosted).ToString("MM/dd/yyyy 00:00:00") + "' where UserRefNo = '" + pPostingModel.cpRefID + "'";
                    bool sqlUpdate = DBUtils.ExecuteSQL(sqlupdate);
                    string Memo = DBUtils.executeSqlGetSingle("Select Memo from GLReferences where GLReferenceID='" + pPostingModel.cpRefID + "'");
                    pPostingModel.Memo = Memo;
                    pPostingModel.isPosted = false;
                    pPostingModel = PostingUtils.ProceedJournalEntry(pPostingModel);
                    if (pPostingModel.isPosted)
                    {
                        bool sqlDelEntry = DBUtils.ExecuteSQL("Delete from GLReferencesAttachment where IsDeleted = 1");

                        if (files != null)
                        {
                            SqlCommand InsertCommand = null;
                            SqlConnection con = Common.getConnection();
                            foreach (var file in files)
                            {
                                var fileName = Path.GetFileName(file.FileName);
                                var physicalPath = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/"), fileName);
                                file.SaveAs(physicalPath);
                                string insertAttachment = "Insert into GLReferencesAttachment (GLReferenceID,FileName,AddedBy,AddedDate) Values (@GLReferenceID,@FileName,@AddedBy,@AddedDate)";
                                ////SqlCommand InsertCommand = null;
                                ////using (SqlConnection con = Common.getConnection())
                                ////{
                                try
                                {
                                    InsertCommand = con.CreateCommand();
                                    InsertCommand.Parameters.AddWithValue("@GLReferenceID", pPostingModel.GLReferenceID);
                                    InsertCommand.Parameters.AddWithValue("@FileName", fileName);
                                    InsertCommand.Parameters.AddWithValue("@AddedBy", profile.Id);
                                    InsertCommand.Parameters.AddWithValue("@AddedDate", Common.toDateTime(SysPrefs.PostingDate));
                                    InsertCommand.CommandText = insertAttachment;
                                    InsertCommand.ExecuteNonQuery();
                                    InsertCommand.Parameters.Clear();
                                }
                                catch (Exception ex)
                                {
                                    pPostingModel.ErrMessage = Common.GetAlertMessage(0, "Error Occured" + ex.ToString());
                                }
                                ////}
                            }

                            con.Close();

                        }
                        pPostingModel.ErrMessage = Common.GetAlertMessage(0, pPostingModel.ErrMessage);
                        return RedirectToAction("ShowMessage", "GLGeneralJournal", new { cpUserRefNo = Security.EncryptQueryString(Common.toString(pPostingModel.cpRefID)) });
                    }
                    else
                    {
                        bool sqlUpdateEntry = DBUtils.ExecuteSQL("Update GLReferencesAttachment set IsDeleted = 0");
                        pPostingModel.ErrMessage = Common.GetAlertMessage(1, pPostingModel.ErrMessage);
                    }
                }
                else
                {
                    pPostingModel.isPosted = true;
                    pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Posting date is not between financial year start date and financial year end date.");
                }
            }
            else
            {
                pPostingModel.isPosted = true;
                pPostingModel.ErrMessage = Common.GetAlertMessage(1, "Invalid pin code provided.");

            }
            return View("~/Views/Postings/GeneralJournal.cshtml", pPostingModel);
        }

        public JsonResult GetDebitCash(string Value)
        {
            AccountDetails objAccount = new AccountDetails();
            objAccount = PostingUtils.GetDebitCash(Value);
            return Json(objAccount);
        }

        public JsonResult Edit(int Id)
        {
            EditViewModel editModel = new EditViewModel();
            editModel = PostingUtils.Edit(Id);
            return Json(editModel);
        }

        public ActionResult Delete(int Id, string cpUserRefNo)
        {
            DeleteViewModel returnModel = new DeleteViewModel();
            returnModel = PostingUtils.Delete(Id, cpUserRefNo);
            return Json(returnModel, JsonRequestBehavior.AllowGet);
        }

        public ActionResult deleteAttachment(int Id, string cpUserRefNo)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            DeleteViewModel returnModel = new DeleteViewModel();
            string attachmentID = Id.ToString();
            if (attachmentID.Trim() != "")
            {
                try
                {
                    string sqlDelEntry = "Update GLReferencesAttachment set IsDeleted = 1 where AttachmentID=" + Common.toString(attachmentID);
                    bool IsDelete = DBUtils.ExecuteSQL(sqlDelEntry);
                    if (IsDelete)
                    {
                        returnModel.isPosted = true;
                        //returnModel.HtmlTable = BuildDynamicTable(Common.toString(cpUserRefNo));
                        string getFileName = "Select * from GLReferencesAttachment where GLReferenceID='" + cpUserRefNo + "' and IsDeleted = 0";
                        DataTable dtFileName = DBUtils.GetDataTable(getFileName);
                        if (dtFileName != null && dtFileName.Rows.Count > 0)
                        {
                            foreach (DataRow drFileName in dtFileName.Rows)
                            {
                                // pGLTransViewModel.Attachments += Common.toString(drFileName["FileName"]) + ", ";
                                returnModel.Attachments += "<a href ='/Uploads/" + SysPrefs.SubmissionFolder + "/Vouchers/" + drFileName["FileName"].ToString() + "', new { target = '_blank' }>" + Common.toString(drFileName["FileName"]) + "</a> <a href='javascript: deleteAttachment(" + Common.toString(drFileName["AttachmentID"]) + ")'><i class='fa fa-remove'></i></a> " + " ,";
                            }
                        }
                    }
                }
                catch
                {
                }
            }
            return Json(returnModel, JsonRequestBehavior.AllowGet);
        }


        #region Update Buyer
        public ActionResult UpdateBuyer()
        {
            UpdateBuyerModel model = new UpdateBuyerModel();
            model.hideContent = true;
            model.isPosted = false;

            return View("~/Views/Postings/UpdateBuyer.cshtml", model);
        }
        public JsonResult GetNewBuyer([DataSourceRequest] DataSourceRequest request)
        {
            ChartOfAccountsModel obj = new ChartOfAccountsModel();
            var bankAccount = _dbcontext.Query<ChartOfAccountsModel>("select VendorFirstName + ' ('+ VendorPrefix + ')' as AccountTitle, VendorId as AccountCode from Vendors order by VendorPrefix").ToList();
            obj.AccountTitle = "Select New Buyer";
            obj.AccountCode = "";
            bankAccount.Insert(0, obj);
            return Json(bankAccount, JsonRequestBehavior.AllowGet);
        }


        [HttpPost]
        public ActionResult UpdateBuyerView(UpdateBuyerModel postingViewModel)
        {

            if (!string.IsNullOrEmpty(postingViewModel.VoucherNo))
            {
                string mySql = "select distinct GLTransactions.GLReferenceID FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID where AccountTypeiCode ='22PAIDENTRY' and VoucherNumber='" + postingViewModel.VoucherNo + "'";
                string GLReferenceID = Common.toString(DBUtils.executeSqlGetSingle(mySql));
                if (!string.IsNullOrEmpty(GLReferenceID))
                {
                    postingViewModel.cpGLReferenceID = GLReferenceID;
                    postingViewModel.HtmlTable1 = BuildDynamicTable1(Common.toInt(postingViewModel.cpGLReferenceID));
                    postingViewModel.isPosted = true;
                    postingViewModel.isHide = false;
                }
                else
                {
                    postingViewModel.isPosted = false;
                    postingViewModel.hideContent = true;
                    postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Unable to find voucher number or voucher is not in paid.");
                }
            }
            else
            {
                postingViewModel.isPosted = false;
                postingViewModel.hideContent = true;
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Please provide voucher no.");
            }
            return View("~/Views/Postings/UpdateBuyer.cshtml", postingViewModel);
        }
        protected static string BuildDynamicTable1(int cpReferenceID)
        {
            string table = "";
            decimal cpTotalDebit = 0m;
            decimal cpTotalCredit = 0m;
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
                        table += "<td valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";
                        string LC = "0.00";
                        if (!string.IsNullOrEmpty(MyRow["LocalCurrencyAmount"].ToString()))
                        {
                            LC = MyRow["LocalCurrencyAmount"].ToString();
                        }
                        decimal LcAmount = Convert.ToDecimal(LC);
                        if (LcAmount < 0)
                        {
                            table += "<td>&nbsp;</td>";
                            table += "<td valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "</td>";
                            cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                        }
                        else
                        {
                            table += "<td valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
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
            table += "<td colspan='2'></td>";
            table += "<td style='background-color:#1c3a70; color:#FFF;'>Total:</td>";
            table += "<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + DisplayUtils.GetSystemAmountFormat(cpTotalDebit) + "</td>";
            table += "<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + DisplayUtils.GetSystemAmountFormat(cpTotalCredit) + "</td>";
            table += "</tr>";
            table += "</tfoot>";
            table += "</table>";
            return table;
        }
        public ActionResult Submit(UpdateBuyerModel model)
        {
            if (!string.IsNullOrEmpty(model.PinCode))
            {
                if (Common.ValidatePinCode(profile.Id, model.PinCode))
                {
                    if (!string.IsNullOrEmpty(model.cpGLReferenceID))
                    {
                        if (!string.IsNullOrEmpty(model.VendorId))
                        {
                            string NewBuyerAccountCode = "", VendorPrefix = "";
                            string sqlVendors = "select VendorId,AccountCode,VendorPrefix from Vendors where VendorId = '" + model.VendorId + "'";
                            DataTable dtVendors = DBUtils.GetDataTable(sqlVendors);
                            if (dtVendors != null)
                            {
                                if (dtVendors.Rows.Count > 0)
                                {
                                    NewBuyerAccountCode = Common.toString(dtVendors.Rows[0]["AccountCode"]).Trim();
                                    VendorPrefix = Common.toString(dtVendors.Rows[0]["VendorPrefix"]).Trim();
                                }
                            }
                            if (NewBuyerAccountCode != "")
                            {
                                string CurrentBuyerAccount = "", CustomerTransactionID = "";
                                string sqlTrans = "select buyerAccountCode,CustomerTransactionID from CustomerTransactions where PaidGLReferenceID=" + model.cpGLReferenceID + "";
                                DataTable dtCustomerTranx = DBUtils.GetDataTable(sqlTrans);
                                if (dtCustomerTranx != null)
                                {
                                    if (dtCustomerTranx.Rows.Count > 0)
                                    {
                                        CurrentBuyerAccount = Common.toString(dtCustomerTranx.Rows[0]["buyerAccountCode"]).Trim();
                                        CustomerTransactionID = Common.toString(dtCustomerTranx.Rows[0]["CustomerTransactionID"]).Trim();
                                        if (!string.IsNullOrEmpty(CurrentBuyerAccount))
                                        {
                                            if (NewBuyerAccountCode.Trim() != CurrentBuyerAccount.Trim())
                                            {
                                                sqlTrans = "select TransactionID from GLTransactions where GLReferenceID=" + model.cpGLReferenceID + " and AccountCode='" + CurrentBuyerAccount + "'";
                                                string TransactionID = DBUtils.executeSqlGetSingle(sqlTrans);
                                                if (!string.IsNullOrEmpty(TransactionID))
                                                {
                                                    string ReferenceID = DBUtils.executeSqlGetSingle("select GLReferenceID from GLReferences where GLReferenceID = " + model.cpGLReferenceID + " and TypeiCode ='22PAIDENTRY'");
                                                    if (!string.IsNullOrEmpty(ReferenceID))
                                                    {
                                                        string sqlUpdate = "Update GLTransactions set AccountCode='" + NewBuyerAccountCode + "' where TransactionID=" + TransactionID;
                                                        DBUtils.ExecuteSQL(sqlUpdate);

                                                        sqlUpdate = "Update CustomerTransactions set BuyerPrefix='" + VendorPrefix + "', BuyerAccountCode='" + NewBuyerAccountCode + "' ,VendorId=" + model.VendorId + " where CustomerTransactionID=" + CustomerTransactionID;
                                                        DBUtils.ExecuteSQL(sqlUpdate);

                                                        model.ErrMessage = Common.GetAlertMessage(0, "Successfully updated New Buyer.<a href='/GlGeneralJournal/UpdateBuyer'>Click here to continue</a>");
                                                        model.isHide = true;
                                                    }
                                                    else
                                                    {
                                                        model.ErrMessage = Common.GetAlertMessage(1, "Sorry! Unable to find transaction details.");

                                                    }
                                                }
                                                else
                                                {
                                                    model.ErrMessage = Common.GetAlertMessage(1, "Sorry! Unable to find transaction details.");
                                                }
                                            }
                                            else
                                            {
                                                model.ErrMessage = Common.GetAlertMessage(1, "Sorry! can not change buyer. Same buyer already assigned.");
                                            }
                                        }
                                        else
                                        {
                                            model.ErrMessage = Common.GetAlertMessage(1, "error in existing buyer info.");
                                        }
                                    }
                                    else
                                    {
                                        model.ErrMessage = Common.GetAlertMessage(1, "error in existing buyer info.");
                                    }
                                }
                                else
                                {
                                    model.ErrMessage = Common.GetAlertMessage(1, "error in existing buyer info.");
                                }
                            }
                            else
                            {
                                model.ErrMessage = Common.GetAlertMessage(1, "New buyer error: account not setup.");
                            }
                            //update GLTransactions set AccountCode= pAccountCode where GLReferenceID=model.cpGLReferenceID

                            //update CustomerTransactions set BuyerPrefix=model.NewBuyerAccount,BuyerAccountCode= pAccountCode where PaidGLReferenceID=model.cpGLReferenceID
                        }
                        else
                        {
                            model.ErrMessage = Common.GetAlertMessage(1, "Please select new buyer.");
                        }
                    }
                    else
                    {
                        model.ErrMessage = Common.GetAlertMessage(1, "Invalid Voucher #.");
                    }
                }
                else
                {
                    model.ErrMessage = Common.GetAlertMessage(1, "Invalid PinCode.");
                }
            }
            else
            {
                model.ErrMessage = Common.GetAlertMessage(1, "Please provide PinCode.");
            }
            model.HtmlTable1 = BuildDynamicTable1(Common.toInt(model.cpGLReferenceID));
            model.isPosted = true;
            return View("~/Views/Postings/UpdateBuyer.cshtml", model);
        }
        #endregion


    }
}