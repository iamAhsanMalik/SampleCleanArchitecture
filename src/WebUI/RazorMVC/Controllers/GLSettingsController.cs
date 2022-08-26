using DHAAccounts.Models;
using Kendo.Mvc.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Mvc;

namespace Accounting.Controllers
{
    public class GLSettingsController : Controller
    {
        ApplicationUser Profile = ApplicationUser.GetUserProfile();
        private AppDbContext _dbcontext = new AppDbContext();
        //// private AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());
        // GET: GLSettings
        public ActionResult Index()
        {
            GLSettingsViewModel gLSettingsViewModel = new GLSettingsViewModel();
            gLSettingsViewModel.DefaultCurrency = Common.GetSysSettings("DefaultCurrency");
            gLSettingsViewModel.AgentBaseAccountCode = Common.GetSysSettings("AgentGLAccount").Replace("\r\n", "");
            gLSettingsViewModel.AgentBaseAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.AgentBaseAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.BuyerBaseAccountCode = Common.GetSysSettings("BuyerGLAccount").Replace("\r\n", "");
            gLSettingsViewModel.BuyerBaseAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.BuyerBaseAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.BankBaseAccountCode = Common.GetSysSettings("BankGLAccount").Replace("\r\n", "");
            gLSettingsViewModel.BankBaseAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.BankBaseAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.UnDepositedAccountCode = Common.GetSysSettings("UnDepositedGLAccount").Replace("\r\n", "");
            gLSettingsViewModel.UnDepositedAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.UnDepositedAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.HoldingAccountCode = Common.GetSysSettings("TransactionHoldingAccount").Replace("\r\n", "");
            gLSettingsViewModel.HoldingAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.HoldingAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.UnearnedServiceFeeAccountCode = Common.GetSysSettings("UnEarnedServiceFeeAccount").Replace("\r\n", "");
            gLSettingsViewModel.UnearnedServiceFeeAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.UnearnedServiceFeeAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.ExchangeVariancesAccountCode = Common.GetSysSettings("DefaultExchangeVariancesAccount").Replace("\r\n", "");
            gLSettingsViewModel.ExchangeVariancesAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.ExchangeVariancesAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.BankChargesAccountCode = Common.GetSysSettings("DefaultBankChargesAccount").Replace("\r\n", "");
            gLSettingsViewModel.BankChargesAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.BankChargesAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.TransactionAdminCommissionAccountCode = Common.GetSysSettings("TransactionAdminCommissionAccount").Replace("\r\n", "");
            gLSettingsViewModel.TransactionAdminCommissionAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.TransactionAdminCommissionAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.TransactionAgentCommissionAccountCode = Common.GetSysSettings("TransactionAgentCommissionAccount").Replace("\r\n", "");
            gLSettingsViewModel.TransactionAgentCommissionAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.TransactionAgentCommissionAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.TranslationGainLossAccountCode = Common.GetSysSettings("TranslationGainLossAccount").Replace("\r\n", "");
            gLSettingsViewModel.TranslationGainLossAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.TranslationGainLossAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossAdministrativeExpensesAccountCode = Common.GetSysSettings("ProfileNLossAdministrativeExpenses").Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossAdministrativeExpensesAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.ProfileNLossAdministrativeExpensesAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossCostOfSalesAccountCode = Common.GetSysSettings("ProfileNLossCostOfSales").Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossCostOfSalesAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.ProfileNLossCostOfSalesAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossIncomeAccountCode = Common.GetSysSettings("ProfileNLossIncomeAccount").Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossIncomeAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.ProfileNLossIncomeAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossOtherExpensesAccountCode = Common.GetSysSettings("ProfileNLossOtherExpenses").Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossOtherExpensesAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.ProfileNLossOtherExpensesAccountCode).Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossOtherIncomeAccountCode = Common.GetSysSettings("ProfileNLossOtherIncome").Replace("\r\n", "");
            gLSettingsViewModel.ProfileNLossOtherIncomeAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.ProfileNLossOtherIncomeAccountCode).Replace("\r\n", "");

            gLSettingsViewModel.BSNonCurrentAssetsAccountCode = Common.GetSysSettings("BSNonCurrentAssetsAccount").Replace("\r\n", "");
            gLSettingsViewModel.BSNonCurrentAssetsAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.BSNonCurrentAssetsAccountCode).Replace("\r\n", "");

            gLSettingsViewModel.BSCurrentAssetsAccountCode = Common.GetSysSettings("BSCurrentAssetsAccount").Replace("\r\n", "");
            gLSettingsViewModel.BSCurrentAssetsAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.BSCurrentAssetsAccountCode).Replace("\r\n", "");

            gLSettingsViewModel.BSCurrentLiabilitesAccountCode = Common.GetSysSettings("BSCurrentLiabilitesAccount").Replace("\r\n", "");
            gLSettingsViewModel.BSCurrentLiabilitesAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.BSCurrentLiabilitesAccountCode).Replace("\r\n", "");

            gLSettingsViewModel.BSNonCurrentliabilitesAccountCode = Common.GetSysSettings("BSNonCurrentliabilitesAccount").Replace("\r\n", "");
            gLSettingsViewModel.BSNonCurrentliabilitesAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.BSNonCurrentliabilitesAccountCode).Replace("\r\n", "");

            gLSettingsViewModel.BSCapitalNReservesAccountCode = Common.GetSysSettings("BSCapitalNReservesAccount").Replace("\r\n", "");
            gLSettingsViewModel.BSCapitalNReservesAccountName = Common.getGLAccountNameCodeByCode(gLSettingsViewModel.BSCapitalNReservesAccountCode).Replace("\r\n", "");

            gLSettingsViewModel.PostingStartDate = Common.toDateTime(Common.GetSysSettings("PostingStartDate").Replace("\r\n", ""));

            if (Common.GetSysSettings("AutoAuthorizeHold") == "1")
            {
                gLSettingsViewModel.AutoAuthorizedHoldingRB = "1";
            }
            else
            {
                gLSettingsViewModel.AutoAuthorizedHoldingRB = "0";
            }

            if (Common.GetSysSettings("AutoAuthorizePaid") == "1")
            {
                gLSettingsViewModel.AutoAuthorizedPaidRB = "1";
            }
            else
            {
                gLSettingsViewModel.AutoAuthorizedPaidRB = "0";
            }

            if (Common.GetSysSettings("AutoAuthorizeCancelled") == "1")
            {
                gLSettingsViewModel.AutoAuthorizedCancelRB = "1";
            }
            else
            {
                gLSettingsViewModel.AutoAuthorizedCancelRB = "0";
            }

            if (Common.GetSysSettings("AutoAuthorizeDeposit") == "1")
            {
                gLSettingsViewModel.AutoAuthorizeDepositRB = "1";
            }
            else
            {
                gLSettingsViewModel.AutoAuthorizeDepositRB = "0";
            }
            if (Common.GetSysSettings("AutoTransferAuthorizationBSU") == "1")
            {
                gLSettingsViewModel.AutoTransferAuthorizationBSU = "1";
            }
            else
            {
                gLSettingsViewModel.AutoTransferAuthorizationBSU = "0";
            }

            if (Common.GetSysSettings("ImportFileHasPostingDate") == "1")
            {
                gLSettingsViewModel.ImportFileHasPostingDate = "1";
            }
            else
            {
                gLSettingsViewModel.ImportFileHasPostingDate = "0";
            }
            ViewBag.inlineDefault = GetChartOfAccountClassicification();
            ViewBag.inlineDefaultHeads = GetChartOfAccountClassicificationWithHeads();
            DataTable dtFinancailDate = DBUtils.GetDataTable("Select * from FinancialYear where Active='true'");
            if (dtFinancailDate != null && dtFinancailDate.Rows.Count > 0)
            {
                gLSettingsViewModel.FinancialYearFromDate = Common.toDateTime(dtFinancailDate.Rows[0]["FinancialYearFromDate"]);
                gLSettingsViewModel.FinancialYearToDate = Common.toDateTime(dtFinancailDate.Rows[0]["FinancialYearToDate"]);
                gLSettingsViewModel.FinancialYearDatePostedBy = Common.toString(dtFinancailDate.Rows[0]["AddedBy"]);
                gLSettingsViewModel.UserEmail = Profile.Email;
            }
            return View("", gLSettingsViewModel);
        }
        private IEnumerable<DropDownTreeItemModel> GetChartOfAccountClassicification()
        {
            List<DropDownTreeItemModel> inlineDefault = new List<DropDownTreeItemModel>();
            inlineDefault = GetChartOfAccountItems("0");
            return inlineDefault;
        }
        private List<DropDownTreeItemModel> GetChartOfAccountItems(string pAccountCode)
        {
            List<DropDownTreeItemModel> pList = new List<DropDownTreeItemModel>();
            if (pAccountCode != "")
            {
                string sqlAccount = "select * from GLChartOfAccounts where isHead=0 And ParentAccountCode='" + pAccountCode + "'";
                DataTable dtAccount = DBUtils.GetDataTable(sqlAccount);
                foreach (DataRow dr in dtAccount.Rows)
                {
                    DropDownTreeItemModel item = new DropDownTreeItemModel();
                    item.Text = dr["AccountName"].ToString().Replace("\r\n", "");
                    item.Id = dr["AccountCode"].ToString().Replace("\r\n", "");
                    item.Expanded = true;
                    item.Items = GetChartOfAccountItems(dr["AccountCode"].ToString());
                    pList.Add(item);
                }
            }
            return pList;
        }
        private IEnumerable<DropDownTreeItemModel> GetChartOfAccountClassicificationWithHeads()
        {
            List<DropDownTreeItemModel> inlineDefault = new List<DropDownTreeItemModel>();
            inlineDefault = GetChartOfAccountItemsWithHeads("0");
            return inlineDefault;
        }
        private List<DropDownTreeItemModel> GetChartOfAccountItemsWithHeads(string pAccountCode)
        {
            List<DropDownTreeItemModel> pList = new List<DropDownTreeItemModel>();
            if (pAccountCode != "")
            {
                string sqlAccount = "select * from GLChartOfAccounts where ParentAccountCode='" + pAccountCode + "'";
                DataTable dtAccount = DBUtils.GetDataTable(sqlAccount);
                foreach (DataRow dr in dtAccount.Rows)
                {
                    DropDownTreeItemModel item = new DropDownTreeItemModel();
                    item.Text = dr["AccountName"].ToString().Replace("\r\n", "");
                    item.Id = dr["AccountCode"].ToString().Replace("\r\n", "");
                    item.Expanded = true;
                    item.Items = GetChartOfAccountItemsWithHeads(dr["AccountCode"].ToString());
                    pList.Add(item);
                }
            }
            return pList;
        }
        public JsonResult LocalCurrencyGLAccount([DataSourceRequest] DataSourceRequest request)
        {
            // SELECT CurrencyISOCode,  (CurrencyISOCode + ' - ' + CurrencyName) as CurrencyName FROM ListCurrencies where CurrencyStatus = '1' order by CurrencyISOCode", "Select Currency
            //  iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();

            //var ListCurrencies = db.ListCurrencies.Where(x => x.CurrencyStatus == true).OrderBy(x => x.CurrencyISOCode).Select(c => new
            //{
            //    CurrencyName = c.CurrencyISOCode.ToString() + " - " + c.CurrencyName.ToString(),
            //    CurrencyISOCode = c.CurrencyISOCode.ToString()

            //}).ToList();

            //  CurrencyModel obj = new CurrencyModel();
            var ListCurrencies = _dbcontext.Query<CurrencyModel>("select  CurrencyName,CurrencyISOCode from ListCurrencies where CurrencyStatus=1 order by  CurrencyISOCode").ToList();
            // ListCurrencies.Add(new { CurrencyISOCode = "", CurrencyName = "Select Currency" });
            return Json(ListCurrencies, JsonRequestBehavior.AllowGet);
        }
        public JsonResult pUpdate(GLSettingsViewModel pGLSettingViewModel)
        {
            string ErrorMessage = "";
            pGLSettingViewModel.isPosted = true;
            if (Common.ValidatePinCode(Profile.Id, pGLSettingViewModel.PinCode))
            {
                //if (pGLSettingViewModel.DefaultCurrency != "" && pGLSettingViewModel.DefaultCurrency != null)
                //{
                //    Common.SetSysSettings("DefaultCurrency", pGLSettingViewModel.DefaultCurrency);
                //}
                //else
                //{
                //    pGLSettingViewModel.message = Common.GetAlertMessage(1, "Please select default currency");
                //}

                /* isHead Validation */
                bool isHeadValid = true;
                if (Common.toString(pGLSettingViewModel.ExchangeVariancesAccountCode).Trim() != "")
                {
                    if (!PostingUtils.isHeadAccountCode(Common.toString(pGLSettingViewModel.ExchangeVariancesAccountCode)))
                    {
                        isHeadValid = false;
                        ErrorMessage += " ,ExchangeVariancesAccount";
                    }

                }
                if (Common.toString(pGLSettingViewModel.TranslationGainLossAccountCode).Trim() != "")
                {
                    if (!PostingUtils.isHeadAccountCode(Common.toString(pGLSettingViewModel.TranslationGainLossAccountCode)))
                    {
                        isHeadValid = false;
                        ErrorMessage += " ,TranslationGainLossAccount";
                    }

                }
                if (Common.toString(pGLSettingViewModel.TransactionAgentCommissionAccountCode).Trim() != "")
                {
                    if (!PostingUtils.isHeadAccountCode(Common.toString(pGLSettingViewModel.TransactionAgentCommissionAccountCode)))
                    {
                        isHeadValid = false;
                        ErrorMessage += " ,TransactionAgentCommissionAccount";
                    }

                }
                if (Common.toString(pGLSettingViewModel.UnearnedServiceFeeAccountCode).Trim() != "")
                {
                    if (!PostingUtils.isHeadAccountCode(Common.toString(pGLSettingViewModel.UnearnedServiceFeeAccountCode)))
                    {
                        isHeadValid = false;
                        ErrorMessage += " ,UnearnedServiceFeeAccount";
                    }

                }
                if (Common.toString(pGLSettingViewModel.TransactionAdminCommissionAccountCode).Trim() != "")
                {
                    if (!PostingUtils.isHeadAccountCode(Common.toString(pGLSettingViewModel.TransactionAdminCommissionAccountCode)))
                    {
                        isHeadValid = false;
                        ErrorMessage += " ,TransactionAdminCommissionAccount";
                    }

                }
                if (Common.toString(pGLSettingViewModel.BankChargesAccountCode).Trim() != "")
                {
                    if (!PostingUtils.isHeadAccountCode(Common.toString(pGLSettingViewModel.BankChargesAccountCode)))
                    {
                        isHeadValid = false;
                        ErrorMessage += " ,BankChargesAccount";
                    }
                }
                if (Common.toString(pGLSettingViewModel.UnDepositedAccountCode).Trim() != "")
                {
                    if (!PostingUtils.isHeadAccountCode(Common.toString(pGLSettingViewModel.UnDepositedAccountCode)))
                    {
                        isHeadValid = false;
                        ErrorMessage += " ,UnDepositedAccount";
                    }
                }

                if (isHeadValid)
                {
                    if (Common.toDateTime(pGLSettingViewModel.FinancialYearFromDate) != null && Common.toDateTime(pGLSettingViewModel.FinancialYearToDate) != null)
                    {
                        using (SqlConnection con = Common.getConnection())
                        {
                            SqlCommand updateCommand = con.CreateCommand();
                            try
                            {
                                string updateFinancailYear = "Update FinancialYear set FinancialYearFromDate='" + Common.toDateTime(pGLSettingViewModel.FinancialYearFromDate).ToString("yyyy/MM/dd 00:00:00") + "' , FinancialYearToDate='" + Common.toDateTime(pGLSettingViewModel.FinancialYearToDate).ToString("yyyy/MM/dd 00:00:00") + "' where Active='true'";
                                updateCommand.CommandText = updateFinancailYear;
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
                    }

                    if (Common.toDateTime(pGLSettingViewModel.PostingStartDate) != null)
                    {
                        if (Common.isDate(pGLSettingViewModel.PostingStartDate))
                        {
                            if (Common.toDateTime(pGLSettingViewModel.PostingStartDate) < Common.toDateTime(SysPrefs.PostingDate))
                            {
                                Common.SetSysSettings("PostingStartDate", Common.toString(pGLSettingViewModel.PostingStartDate));
                            }
                        }
                    }


                    Common.SetSysSettings("AgentGLAccount", Common.toString(pGLSettingViewModel.AgentBaseAccountCode));
                    Common.SetSysSettings("BuyerGLAccount", Common.toString(pGLSettingViewModel.BuyerBaseAccountCode));
                    Common.SetSysSettings("BankGLAccount", Common.toString(pGLSettingViewModel.BankBaseAccountCode));
                    Common.SetSysSettings("DefaultExchangeVariancesAccount", Common.toString(pGLSettingViewModel.ExchangeVariancesAccountCode));
                    Common.SetSysSettings("DefaultBankChargesAccount", Common.toString(pGLSettingViewModel.BankChargesAccountCode));
                    Common.SetSysSettings("UnDepositedGLAccount", Common.toString(pGLSettingViewModel.UnDepositedAccountCode));
                    Common.SetSysSettings("TransactionHoldingAccount", Common.toString(pGLSettingViewModel.HoldingAccountCode));
                    //Common.SetSysSettings("TransactionCommissionAccount", Common.toString(pGLSettingViewModel.CommissionAccountCode));
                    Common.SetSysSettings("UnEarnedServiceFeeAccount", Common.toString(pGLSettingViewModel.UnearnedServiceFeeAccountCode));
                    //Common.SetSysSettings("DefaultInventoryAccount", Common.toString(pGLSettingViewModel.InventoryAccountCode));
                    Common.SetSysSettings("TransactionAdminCommissionAccount", Common.toString(pGLSettingViewModel.TransactionAdminCommissionAccountCode));
                    Common.SetSysSettings("TransactionAgentCommissionAccount", Common.toString(pGLSettingViewModel.TransactionAgentCommissionAccountCode));
                    Common.SetSysSettings("TranslationGainLossAccount", Common.toString(pGLSettingViewModel.TranslationGainLossAccountCode));

                    Common.SetSysSettings("ProfileNLossAdministrativeExpenses", Common.toString(pGLSettingViewModel.ProfileNLossAdministrativeExpensesAccountCode));
                    Common.SetSysSettings("ProfileNLossCostOfSales", Common.toString(pGLSettingViewModel.ProfileNLossCostOfSalesAccountCode));
                    Common.SetSysSettings("ProfileNLossIncomeAccount", Common.toString(pGLSettingViewModel.ProfileNLossIncomeAccountCode));
                    Common.SetSysSettings("ProfileNLossOtherExpenses", Common.toString(pGLSettingViewModel.ProfileNLossOtherExpensesAccountCode));
                    Common.SetSysSettings("ProfileNLossOtherIncome", Common.toString(pGLSettingViewModel.ProfileNLossOtherIncomeAccountCode));

                    Common.SetSysSettings("BSNonCurrentAssetsAccount", Common.toString(pGLSettingViewModel.BSNonCurrentAssetsAccountCode));
                    Common.SetSysSettings("BSCurrentAssetsAccount", Common.toString(pGLSettingViewModel.BSCurrentAssetsAccountCode));
                    Common.SetSysSettings("BSCurrentLiabilitesAccount", Common.toString(pGLSettingViewModel.BSCurrentLiabilitesAccountCode));
                    Common.SetSysSettings("BSNonCurrentliabilitesAccount", Common.toString(pGLSettingViewModel.BSNonCurrentliabilitesAccountCode));
                    Common.SetSysSettings("BSCapitalNReservesAccount", Common.toString(pGLSettingViewModel.BSCapitalNReservesAccountCode));

                    if (Common.toString(pGLSettingViewModel.AutoAuthorizedHoldingRB) == "True")
                    {
                        Common.SetSysSettings("AutoAuthorizeHold", "1");
                    }
                    else
                    {
                        Common.SetSysSettings("AutoAuthorizeHold", "0");
                    }

                    if (Common.toString(pGLSettingViewModel.AutoAuthorizedPaidRB) == "True")
                    {
                        Common.SetSysSettings("AutoAuthorizePaid", "1");
                    }
                    else
                    {
                        Common.SetSysSettings("AutoAuthorizePaid", "0");
                    }

                    if (Common.toString(pGLSettingViewModel.AutoAuthorizedCancelRB) == "True")
                    {
                        Common.SetSysSettings("AutoAuthorizeCancelled", "1");
                    }
                    else
                    {
                        Common.SetSysSettings("AutoAuthorizeCancelled", "0");
                    }

                    if (Common.toString(pGLSettingViewModel.AutoAuthorizeDepositRB) == "True")
                    {
                        Common.SetSysSettings("AutoAuthorizeDeposit", "1");
                    }
                    else
                    {
                        Common.SetSysSettings("AutoAuthorizeDeposit", "0");
                    }

                    if (Common.toString(pGLSettingViewModel.AutoTransferAuthorizationBSU) == "True")
                    {
                        Common.SetSysSettings("AutoTransferAuthorizationBSU", "1");
                    }
                    else
                    {
                        Common.SetSysSettings("AutoTransferAuthorizationBSU", "0");
                    }

                    if (Common.toString(pGLSettingViewModel.ImportFileHasPostingDate) == "True")
                    {
                        Common.SetSysSettings("ImportFileHasPostingDate", "1");
                    }
                    else
                    {
                        Common.SetSysSettings("ImportFileHasPostingDate", "0");
                    }

                    if (pGLSettingViewModel.isPosted)
                    {
                        SysPrefs.Reload();
                        pGLSettingViewModel.message = Common.GetAlertMessage(0, "&nbsp &nbsp &nbsp Settings successfully saved.");
                    }
                    else
                    {
                        pGLSettingViewModel.message = Common.GetAlertMessage(1, "Error Occured");
                    }
                }
                else
                {
                    ErrorMessage = ErrorMessage.TrimStart(',');
                    ErrorMessage += "      these are not Head Account(s).";
                    //  pGLSettingViewModel.message = Common.GetAlertMessage(1, "Please setup all accounts in [Default Accounts] section");
                    pGLSettingViewModel.message = Common.GetAlertMessage(1, ErrorMessage);
                }
            }
            else
            {
                pGLSettingViewModel.message = Common.GetAlertMessage(1, "Invalid pin code provided.");
            }
            return Json(pGLSettingViewModel, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public ActionResult Update(GLSettingsViewModel pGLSettingViewModel)
        {
            pGLSettingViewModel.isPosted = true;
            if (Common.ValidatePinCode(Profile.Id, pGLSettingViewModel.PinCode))
            {
                //if (pGLSettingViewModel.DefaultCurrency != "" && pGLSettingViewModel.DefaultCurrency != null)
                //{
                //    Common.SetSysSettings("DefaultCurrency", pGLSettingViewModel.DefaultCurrency);
                //    pGLSettingViewModel.message = Common.GetAlertMessage(0, pGLSettingViewModel.message);
                //}
                //else
                //{
                //    pGLSettingViewModel.message = Common.GetAlertMessage(1, pGLSettingViewModel.message);
                //}

                Common.SetSysSettings("AgentGLAccount", pGLSettingViewModel.AgentBaseAccountCode);
                Common.SetSysSettings("BuyerGLAccount", pGLSettingViewModel.BuyerBaseAccountCode);
                Common.SetSysSettings("BankGLAccount", pGLSettingViewModel.BankBaseAccountCode);
                Common.SetSysSettings("DefaultExchangeVariancesAccount", pGLSettingViewModel.ExchangeVariancesAccountCode);
                Common.SetSysSettings("DefaultBankChargesAccount", pGLSettingViewModel.BankChargesAccountCode);
                Common.SetSysSettings("UnDepositedGLAccount", pGLSettingViewModel.UnDepositedAccountCode);
                Common.SetSysSettings("TransactionHoldingAccount", pGLSettingViewModel.HoldingAccountCode);
                Common.SetSysSettings("TransactionCommissionAccount", pGLSettingViewModel.CommissionAccountCode);
                Common.SetSysSettings("UnEarnedServiceFeeAccount", pGLSettingViewModel.UnearnedServiceFeeAccountCode);
                Common.SetSysSettings("DefaultInventoryAccount", pGLSettingViewModel.InventoryAccountCode);

                if (pGLSettingViewModel.isPosted)
                {
                    pGLSettingViewModel.message = Common.GetAlertMessage(0, "Settings successfully saved.");
                }
                else
                {
                    pGLSettingViewModel.message = Common.GetAlertMessage(1, "Error Occured");
                }
            }
            else
            {
                pGLSettingViewModel.message = Common.GetAlertMessage(1, "Invalid pin code provided.");
            }
            ViewBag.inlineDefault = GetChartOfAccountClassicification();
            ViewBag.inlineDefaultHeads = GetChartOfAccountClassicificationWithHeads();
            return View("~/Views/GLSettings/Index.cshtml", pGLSettingViewModel);
        }

        public ActionResult PostingDate()
        {
            PostingDateViewModel pPostingDateViewModel = new PostingDateViewModel();
            pPostingDateViewModel.MaxDate = DateTime.Now;
            string abc = Common.GetSysSettings("PostingDate");
            if (Common.GetSysSettings("PostingDate") != "")
            {
                pPostingDateViewModel.PostingDate = Common.toDateTime(Common.GetSysSettings("PostingDate"));
                //pPostingDateViewModel.PostingDate = DateTime.Now;
            }
            else
            {
                pPostingDateViewModel.PostingDate = DateTime.Now;
            }
            return View("~/Views/GLSettings/PostingDate.cshtml", pPostingDateViewModel);
        }
        [HttpPost]
        public ActionResult PostingDate(PostingDateViewModel pPostingDateViewModel)
        {
            pPostingDateViewModel.isPosted = false;
            //PostingDateViewModel pPostingDateViewModel = new PostingDateViewModel();
            if (Common.isDate(pPostingDateViewModel.PostingDate))
            {
                var begin = Common.GetFiscalYearBeginForDate(pPostingDateViewModel.PostingDate);
                if (begin > pPostingDateViewModel.PostingDate)
                {
                    //phMessage.Visible = true;
                    pPostingDateViewModel.message = pPostingDateViewModel.message + (pPostingDateViewModel.message == "" ? "" : "Invalid Posting date . Post Date Must be greater or equal to Financial year date ");
                }
                else
                {
                    if (Profile != null)
                    {
                        string postingStartDate = DBUtils.executeSqlGetSingle("SELECT KeyValue FROM SysSettings where KeyName='PostingStartDate'");
                        if (Common.toString(postingStartDate) != "")
                        {
                            if (Common.isDate(postingStartDate))
                            {
                                if (Common.toDateTime(postingStartDate) > pPostingDateViewModel.PostingDate)
                                {
                                    pPostingDateViewModel.message = pPostingDateViewModel.message + (pPostingDateViewModel.message == "" ? "" : "Invalid Posting date . Post Date Must be greater or equal to posting start date");
                                }
                                else
                                {
                                    if (Common.ValidatePinCode(Profile.Id, pPostingDateViewModel.PinCode))
                                    {
                                        DateTime PostingDate = Common.toDateTime(pPostingDateViewModel.PostingDate);
                                        Common.SetSysSettings("PostingDate", pPostingDateViewModel.PostingDate.ToShortDateString());
                                        SysPrefs.Reload();
                                        pPostingDateViewModel.isPosted = true;
                                        pPostingDateViewModel.message = pPostingDateViewModel.message + (pPostingDateViewModel.message == "" ? "" : "Posting Date updated successfully.");
                                    }
                                    else
                                    {
                                        pPostingDateViewModel.message = pPostingDateViewModel.message + (pPostingDateViewModel.message == "" ? "" : "Invalid pin code entered.");
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (Common.ValidatePinCode(Profile.Id, pPostingDateViewModel.PinCode))
                            {
                                DateTime PostingDate = Common.toDateTime(pPostingDateViewModel.PostingDate);
                                Common.SetSysSettings("PostingDate", pPostingDateViewModel.PostingDate.ToShortDateString());
                                SysPrefs.Reload();
                                pPostingDateViewModel.isPosted = true;
                                pPostingDateViewModel.message = pPostingDateViewModel.message + (pPostingDateViewModel.message == "" ? "" : "Posting Date updated successfully.");
                            }
                            else
                            {
                                pPostingDateViewModel.message = pPostingDateViewModel.message + (pPostingDateViewModel.message == "" ? "" : "Invalid pin code entered.");
                            }
                        }

                    }
                    else
                    {
                        pPostingDateViewModel.message = pPostingDateViewModel.message + (pPostingDateViewModel.message == "" ? "" : "Current session has been expired.please login again.");
                    }
                }
            }
            else
            {
                //lblMessage.Text = Utils.GetAlertMessage(0, "Changes done successfully");                
                pPostingDateViewModel.message = pPostingDateViewModel.message + (pPostingDateViewModel.message == "" ? "" : "Invalid date entered.");
            }
            if (pPostingDateViewModel.isPosted)
            {
                pPostingDateViewModel.message = Common.GetAlertMessage(0, pPostingDateViewModel.message);
            }
            else
            {
                pPostingDateViewModel.message = Common.GetAlertMessage(1, pPostingDateViewModel.message);
            }
            return View("~/Views/GLSettings/PostingDate.cshtml", pPostingDateViewModel);
            //return Json(pPostingDateViewModel, JsonRequestBehavior.AllowGet);
        }
    }
}
