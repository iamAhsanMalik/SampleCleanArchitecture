using DHAAccounts.Models;
using Kendo.Mvc.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace Accounting.Controllers
{
    public class GLRecptEntryController : Controller
    {
        #region Whole Controller
        //private iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
        ApplicationUser profile = ApplicationUser.GetUserProfile();
        private AppDbContext _dbcontext = new AppDbContext();
        //// private AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());
        const string TypeiCode = "22BANKPMT";

        public List<SelectListItem> Customers { get; private set; }


        // GET: Postings
        //public ActionResult Async_Save(IEnumerable<HttpPostedFileBase> files)
        //{
        //    if (files != null)
        //    {
        //        foreach (var file in files)
        //        {
        //            var fileName = Path.GetFileName(file.FileName);
        //            var physicalPath = Path.Combine(Server.MapPath("~/Uploads/Vouchers/"), fileName);
        //            file.SaveAs(physicalPath);
        //        }
        //    }
        //    return Content("");
        //}
        //public ActionResult Async_Remove(string[] fileNames)
        //{
        //    if (fileNames != null)
        //    {
        //        foreach (var fullName in fileNames)
        //        {
        //            var fileName = Path.GetFileName(fullName);
        //            var physicalPath = Path.Combine(Server.MapPath("~/Uploads/Vouchers/"), fileName);
        //            if (System.IO.File.Exists(physicalPath))
        //            {
        //                System.IO.File.Delete(physicalPath);
        //            }
        //        }
        //    }
        //    return Content("");
        //}

        // GET: GLRecptEntry

        public ActionResult Index()
        {
            return View();
        }
        // GET: CashEntry
        public ActionResult RecptEntry(int? id)
        {

            PostingViewModel postingViewModel = new PostingViewModel();
            Guid guid = Guid.NewGuid();
            postingViewModel.cpUserRefNo = guid.ToString();
            Session["cpUserRefNo"] = postingViewModel.cpUserRefNo;
            if (id > 0)
            {
                //AddUpdate()
                FillData(Common.toInt(id));
                postingViewModel.GLReferenceID = Common.toInt(id);
                // GLReferences gLReference = db.GLReferences.Find(Common.toInt(id));
                // postingViewModel.cpUserRefNo = gLReference.ReferenceNo;
                postingViewModel.cpUserRefNo = DBUtils.executeSqlGetSingle("select ReferenceNo from GLReferences where GLReferenceID=" + id + "");
            }
            postingViewModel.table = BuildDynamicTable(postingViewModel.cpUserRefNo);
            return View("~/Views/Postings/RecptEntry.cshtml", postingViewModel);
        }

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
        public ActionResult ProceedRecptEntry(PostingViewModel pPostingModel, IEnumerable<HttpPostedFileBase> files)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            pPostingModel.isPosted = false;
            pPostingModel = PostingUtils.ProceedRecptEntry(pPostingModel);
            if (pPostingModel.isPosted)
            {
                if (files != null)
                {
                    int i = 1;
                    SqlCommand InsertCommand = null;
                    SqlConnection con = Common.getConnection();
                    foreach (var file in files)
                    {
                        var fileName = Path.GetFileName(file.FileName);
                        fileName = pPostingModel.GLReferenceID + "-" + i + "-" + fileName;
                        var physicalPath = Path.Combine(Server.MapPath("~/Uploads/" + SysPrefs.SubmissionFolder + "/Vouchers/"), fileName);
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
                        ////  }
                        i++;
                    }

                    con.Close();
                }
                pPostingModel.ErrMessage = Common.GetAlertMessage(0, pPostingModel.ErrMessage);
                return RedirectToAction("ShowMessage", "GLRecptEntry", new { cpUserRefNo = Security.EncryptQueryString(Common.toString(pPostingModel.GLReferenceID)) });
            }
            else
            {
                pPostingModel.ErrMessage = Common.GetAlertMessage(1, pPostingModel.ErrMessage);
            }
            return View("~/Views/Postings/RecptEntry.cshtml", pPostingModel);
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
        #endregion
        public async Task<ActionResult> GetCustomers()
        {
            var dbHelpers = new DbHelpers();
            var customers = await dbHelpers.PopulateDropDownAsync("select AccountCode as Value, AccountName as Text from GLChartOfAccounts where ParentAccountCode = '113' order by AccountName asc", "AccountName", "AccountCode", "Select Customer");
            return Json(customers, JsonRequestBehavior.AllowGet);
        }
    }
}