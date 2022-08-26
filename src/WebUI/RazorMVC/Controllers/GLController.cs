using DHAAccounts.Models;
using Kendo.Mvc.UI;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace Accounting.Controllers
{
    public class GLController : Controller
    {
        string TypeiCode = "22CANCELENTRY";
        ApplicationUser Profile = ApplicationUser.GetUserProfile();
        private AppDbContext _dbContext = new AppDbContext();
        ////private AppDbContext _dbContext = new AppDbContext(BaseModel.getConnString());
        // GET: GL
        #region Cancel Payment
        public ActionResult CancelPayment()
        {
            CancelPaymentViewModel model = new CancelPaymentViewModel();
            model.hideContent = true;
            return View("~/Views/Postings/CancelPayment.cshtml", model);
        }
        public JsonResult GetGeneralAccounts([DataSourceRequest] DataSourceRequest request)
        {
            //iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
            //var bankAccount = db.GLChartOfAccounts.Where(x => x.isHead == true && x.AccountStatus == true).OrderBy(x => x.AccountName).Select(c => new
            //{
            //    AccountCode = c.AccountCode,
            //    AccountTitle = c.AccountName

            //}).ToList();
            //  bankAccount.Insert(0, new { AccountCode = "", AccountTitle = "Select Charges Account" });
            ChartOfAccountsModel obj = new ChartOfAccountsModel();
            var bankAccount = _dbContext.Query<ChartOfAccountsModel>("select AccountCode as AccountCode, AccountName as AccountTitle from GLChartOfAccounts where isHead=1 and AccountStatus=1 order by  AccountName").ToList();
            obj.AccountTitle = "Select Charges Account";
            obj.AccountCode = "";
            bankAccount.Insert(0, obj);
            return Json(bankAccount, JsonRequestBehavior.AllowGet);
        }
        public ActionResult CancelPaymentView(CancelPaymentViewModel pGLTransViewModel)
        {
            pGLTransViewModel.isPosted = false;
            pGLTransViewModel.hideContent = true;
            if (!string.IsNullOrEmpty(pGLTransViewModel.tranxRef))
            {
                string Records = "";
                if (pGLTransViewModel.Type == "1")
                {
                    Records = "Select CustomerTransactionID,PaymentNo, HoldDate, CancelledGLReferenceID, CancelledDate, SendingCountry,RecevingCountry,Address,City,TransactionID,Phone,Date,Recipient,SenderName,FC_Amount,CancellationReason,(select top 1 CustomerPrefix from Customers where Customers.CustomerID=CustomerTransactions.CustomerID ) as  CustomerPrefix,Status,PaidDate,Pounds,PaidGLReferenceID,HoldGLReferenceID  from CustomerTransactions Where PaymentNo ='" + pGLTransViewModel.tranxRef + "'";
                }
                else if (pGLTransViewModel.Type == "2")
                {
                    Records = "Select CustomerTransactionID,PaymentNo, HoldDate,CancelledGLReferenceID, CancelledDate, SendingCountry,RecevingCountry,Address,City,TransactionID,Phone,Date,Recipient,SenderName,FC_Amount,CancellationReason,(select top 1 CustomerPrefix from Customers where Customers.CustomerID=CustomerTransactions.CustomerID ) as  CustomerPrefix,Status,PaidDate,Pounds,PaidGLReferenceID,HoldGLReferenceID  from CustomerTransactions Where TransactionID='" + pGLTransViewModel.tranxRef + "'";
                }
                DataTable transTable = DBUtils.GetDataTable(Records);
                if (transTable != null)
                {
                    if (transTable.Rows.Count > 0)
                    {
                        DataRow drTransaction = transTable.Rows[0];
                        if (drTransaction != null)
                        {
                            pGLTransViewModel.CustomerTransactionID = Common.toString(drTransaction["CustomerTransactionID"]);
                            pGLTransViewModel.TransactionID = Common.toString(drTransaction["TransactionID"]);
                            if (Common.toString(drTransaction["Status"]).Trim() != "Cancelled")
                            {
                                if (Common.toString(drTransaction["CancelledGLReferenceID"]) == "" || Common.toString(drTransaction["CancelledGLReferenceID"]) == "0")
                                {
                                    string getVoucherNo = "SELECT distinct GLReferences.VoucherNumber, GLTransactions.GLReferenceID FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID where TransactionNumber = '" + Common.toString(drTransaction["PaymentNo"]) + "'";
                                    DataTable dtVouchers = DBUtils.GetDataTable(getVoucherNo);
                                    if (dtVouchers != null && dtVouchers.Rows.Count > 0)
                                    {
                                        foreach (DataRow drVoucher in dtVouchers.Rows)
                                        {
                                            pGLTransViewModel.VoucherNo += Common.toString(drVoucher["VoucherNumber"]) + ", ";
                                            pGLTransViewModel.cpGLReferenceID = Common.toString(drVoucher["GLReferenceID"]);
                                        }
                                    }
                                    string getAccountName = "select AccountName, AccountCode, CurrencyISOCode from GLChartOfAccounts where Prefix = '" + Common.toString(drTransaction["CustomerPrefix"]) + "'";
                                    DataTable getHoldAcc = DBUtils.GetDataTable(getAccountName);
                                    if (getHoldAcc != null && getHoldAcc.Rows.Count > 0)
                                    {
                                        DataRow drRow = getHoldAcc.Rows[0];
                                        pGLTransViewModel.CurrencyISOCode = Common.toString(drRow["CurrencyISOCode"]);
                                        pGLTransViewModel.AccountName = Common.toString(drRow["AccountName"]);
                                        pGLTransViewModel.HoldAccount = Common.toString(drRow["AccountCode"]);

                                        pGLTransViewModel.CurrencyType = DBUtils.executeSqlGetSingle("Select CurrenyType from ListCurrencies Where CurrencyISOCode = '" + pGLTransViewModel.CurrencyISOCode + "'");
                                        pGLTransViewModel.AgentCurrRate = DBUtils.executeSqlGetSingle("Select ExchangeRate from ExchangeRates Where CurrencyISOCode='" + pGLTransViewModel.CurrencyISOCode + "' ");
                                    }
                                    pGLTransViewModel.CustomerPrefix = Common.toString(drTransaction["CustomerPrefix"]);
                                    pGLTransViewModel.Phone = Common.toString(drTransaction["Phone"]);
                                    pGLTransViewModel.PaymentNo = Common.toString(drTransaction["PaymentNo"]);
                                    pGLTransViewModel.InvoiceNo = Common.toString(drTransaction["TransactionID"]);
                                    pGLTransViewModel.Date = Common.toDateTime(drTransaction["Date"]);
                                    pGLTransViewModel.CancelledDate = Common.toDateTime(drTransaction["CancelledDate"]);
                                    pGLTransViewModel.Recipient = Common.toString(drTransaction["Recipient"]) + ", " + Common.toString(drTransaction["Address"] + "-" + Common.toString(drTransaction["City"]));
                                    pGLTransViewModel.SenderName = Common.toString(drTransaction["SenderName"] + "," + drTransaction["SendingCountry"]);
                                    pGLTransViewModel.FC_Amount = Common.toDecimal(drTransaction["FC_Amount"]);
                                    pGLTransViewModel.CancellationReason = Common.toString(drTransaction["CancellationReason"]);
                                    pGLTransViewModel.Status = Common.toString(drTransaction["Status"]);
                                    pGLTransViewModel.PaidDate = Common.toDateTime(drTransaction["PaidDate"]);
                                    pGLTransViewModel.HoldDate = Common.toDateTime(drTransaction["HoldDate"]);
                                    pGLTransViewModel.Pounds = Common.toDecimal(drTransaction["Pounds"]);
                                    pGLTransViewModel.PaidGLReferenceID = Common.toString(drTransaction["PaidGLReferenceID"]);
                                    pGLTransViewModel.HoldGLReferenceID = Common.toString(drTransaction["HoldGLReferenceID"]);
                                    pGLTransViewModel = CancelTranxBuildTable(pGLTransViewModel);
                                    //if (pGLTransViewModel.Status == "Paid")
                                    //{
                                    //    pGLTransViewModel.PaidGLReferenceID = Common.toString(drTransaction["PaidGLReferenceID"]);
                                    //    pGLTransViewModel.HoldGLReferenceID = Common.toString(drTransaction["HoldGLReferenceID"]);
                                    //    pGLTransViewModel = CancelTranxBuildTable(pGLTransViewModel);
                                    //}
                                    //else
                                    //{
                                    //    pGLTransViewModel.HoldGLReferenceID = Common.toString(drTransaction["HoldGLReferenceID"]);
                                    //    pGLTransViewModel = CancelTranxBuildTable(pGLTransViewModel);
                                    //}
                                    pGLTransViewModel.isPosted = true;
                                }
                                else
                                {
                                    pGLTransViewModel.CancellationMessage = "&nbsp &nbsp &nbsp &nbsp <h4>Transaction already cancelled and pending for authorization.</h4>";
                                }
                            }
                            else
                            {
                                pGLTransViewModel.CancellationMessage = "&nbsp &nbsp &nbsp &nbsp <h4>Transaction already cancelled.</h4>";
                            }

                        }
                        else
                        {
                            pGLTransViewModel.ErrMessage = Common.GetAlertMessage(1, "Transaction not found.");
                        }
                    }
                    else
                    {
                        pGLTransViewModel.ErrMessage = Common.GetAlertMessage(1, "No data found against  <b>" + pGLTransViewModel.tranxRef + " </b> Invoice/Payment No.");
                    }

                }
                else
                {
                    pGLTransViewModel.ErrMessage = Common.GetAlertMessage(1, "No data found against  <b>" + pGLTransViewModel.tranxRef + "</b>  Invoice/Payment No.");
                }
            }
            else
            {
                pGLTransViewModel.ErrMessage = Common.GetAlertMessage(1, "Invalid Payment#/Invoice#");
            }
            return View("~/Views/Postings/CancelPayment.cshtml", pGLTransViewModel);
        }
        protected CancelPaymentViewModel CancelTranxBuildTable(CancelPaymentViewModel model)
        {
            string table = "";
            decimal cpTotalDebit = 0m;
            decimal cpTotalCredit = 0m;
            string FC = "0.00";
            string LC = "0.00";
            int iRowNumber = 0;
            decimal LcAmount = 0m;
            decimal FcAmount = 0m;
            string tranxQury = "";
            if (model.Status == "Paid")
            {
                tranxQury = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.PaidGLReferenceID + "' order by LocalCurrencyAmount Desc";
            }
            else
            {
                tranxQury = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' order by LocalCurrencyAmount Desc";
            }
            DataTable objGLTransView = DBUtils.GetDataTable(tranxQury);
            if (objGLTransView != null && objGLTransView.Rows.Count > 0)
            {

                foreach (DataRow MyRow in objGLTransView.Rows)
                {
                    string sqlListCurrencies = "select HoldAccountCode from ListCurrencies where CurrencyISOCode='" + Common.toString(MyRow["ForeignCurrencyISOCode"]).Trim() + "'";
                    string myHoldAccountCode = DBUtils.executeSqlGetSingle(sqlListCurrencies);

                    if (model.Status == "Paid" && myHoldAccountCode == Common.toString(MyRow["AccountCode"]).Trim())
                    {
                        string myAgentAccountCode = "";
                        string myAgentAccountName = "";
                        string sqlGLChartOfAccounts = "select AccountCode,AccountName from  GLChartOfAccounts where prefix='" + model.CustomerPrefix + "'";
                        DataTable dtAgentNameCode = DBUtils.GetDataTable(sqlGLChartOfAccounts);
                        if (dtAgentNameCode != null && dtAgentNameCode.Rows.Count > 0)
                        {
                            myAgentAccountCode = dtAgentNameCode.Rows[0]["AccountCode"].ToString();
                            myAgentAccountName = dtAgentNameCode.Rows[0]["AccountName"].ToString();
                        }
                        string sqlAgentTransaction = "select AccountCode,ForeignCurrencyISOCode, GLReferenceID,LocalCurrencyAmount,ForeignCurrencyAmount,ExchangeRate,BaseAmount,Memo,(Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName from GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode ='" + myAgentAccountCode + "' and LocalCurrencyAmount='" + MyRow["LocalCurrencyAmount"].ToString() + "'";
                        DataTable dtAgentTransaction = DBUtils.GetDataTable(sqlAgentTransaction);
                        if (dtAgentTransaction != null && dtAgentTransaction.Rows.Count > 0)
                        {
                            DataRow drAgentTransaction = dtAgentTransaction.Rows[0];
                            if (!string.IsNullOrEmpty(drAgentTransaction["ForeignCurrencyAmount"].ToString()))
                            {
                                FC = drAgentTransaction["ForeignCurrencyAmount"].ToString();
                            }
                            table += "<tr><td style='padding: 2px' valign='top'><strong>" + drAgentTransaction["AccountCode"].ToString() + "-" + drAgentTransaction["AccountName"].ToString() + "</strong>" + "<br/>" + drAgentTransaction["Memo"].ToString() + "<br/>" + drAgentTransaction["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAgentTransaction["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAgentTransaction["GLReferenceID"].ToString() + "</td>";
                            table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAgentTransaction["ForeignCurrencyISOCode"].ToString() + "</td>";
                            table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";

                            if (!string.IsNullOrEmpty(drAgentTransaction["LocalCurrencyAmount"].ToString()))
                            {
                                LC = drAgentTransaction["LocalCurrencyAmount"].ToString();
                            }
                            LcAmount = -1 * Convert.ToDecimal(LC);

                            if (LcAmount < 0)
                            {
                                table += "<td style='padding: 2px'>&nbsp;</td>";
                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "</td>";
                                cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                                model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                            }
                            else
                            {
                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                table += "<td style='padding: 2px'>&nbsp;</td>";
                                cpTotalDebit = cpTotalDebit + LcAmount;
                                model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                            }
                            table += "</tr>";

                            #region In case of Paid 
                            if (model.Status == "Paid")
                            {
                                #region Admin charges
                                if (Common.toString(SysPrefs.TransactionAdminCommissionAccount) != "" && Common.toString(model.HoldGLReferenceID) != "")
                                {
                                    string sqlAdminCharges = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode='" + SysPrefs.TransactionAdminCommissionAccount + "' order by LocalCurrencyAmount Desc";
                                    DataTable ObjAdminCharges = DBUtils.GetDataTable(sqlAdminCharges);
                                    if (ObjAdminCharges != null && ObjAdminCharges.Rows.Count > 0)
                                    {
                                        iRowNumber = 0;
                                        foreach (DataRow drAdminCharges in ObjAdminCharges.Rows)
                                        {
                                            //Debit Agent account charges
                                            FC = "0.00";
                                            if (!string.IsNullOrEmpty(drAdminCharges["ForeignCurrencyAmount"].ToString()))
                                            {
                                                FC = drAdminCharges["ForeignCurrencyAmount"].ToString();
                                            }
                                            FcAmount = Math.Abs(Convert.ToDecimal(FC));
                                            table += "<tr><td style='padding: 2px' valign='top'><strong>" + drAdminCharges["AccountCode"].ToString() + "-" + drAdminCharges["AccountName"].ToString() + "</strong>" + "<br/>" + drAdminCharges["Memo"].ToString() + "<br/>" + drAdminCharges["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAdminCharges["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAdminCharges["GLReferenceID"].ToString() + "</td>";
                                            table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAdminCharges["ForeignCurrencyISOCode"].ToString() + "</td>";
                                            table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";
                                            LC = "0.00";
                                            if (!string.IsNullOrEmpty(drAdminCharges["LocalCurrencyAmount"].ToString()))
                                            {
                                                LC = drAdminCharges["LocalCurrencyAmount"].ToString();
                                            }
                                            LcAmount = -1 * Convert.ToDecimal(LC);

                                            if (LcAmount < 0)
                                            {
                                                table += "<td style='padding: 2px'>&nbsp;</td>";
                                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "</td>";
                                                cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                                                model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                                            }
                                            else
                                            {
                                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                                table += "<td style='padding: 2px'>&nbsp;</td>";
                                                cpTotalDebit = cpTotalDebit + LcAmount;
                                                model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                                            }
                                            table += "</tr>";
                                            iRowNumber++;

                                            //Credit Agent account charges

                                            FC = "0.00";
                                            if (!string.IsNullOrEmpty(drAdminCharges["ForeignCurrencyAmount"].ToString()))
                                            {
                                                FC = drAdminCharges["ForeignCurrencyAmount"].ToString();
                                            }
                                            FcAmount = Math.Abs(Convert.ToDecimal(FC));
                                            table += "<tr><td style='padding: 2px' valign='top'><strong>" + myAgentAccountCode + "-" + myAgentAccountName + "</strong>" + "<br/>" + drAdminCharges["Memo"].ToString() + "<br/>" + drAdminCharges["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAdminCharges["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAdminCharges["GLReferenceID"].ToString() + "</td>";
                                            table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAdminCharges["ForeignCurrencyISOCode"].ToString() + "</td>";
                                            table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";
                                            LC = "0.00";
                                            if (!string.IsNullOrEmpty(drAdminCharges["LocalCurrencyAmount"].ToString()))
                                            {
                                                LC = drAdminCharges["LocalCurrencyAmount"].ToString();
                                            }
                                            LcAmount = 1 * Convert.ToDecimal(LC);

                                            if (LcAmount < 0)
                                            {
                                                table += "<td style='padding: 2px'>&nbsp;</td>";
                                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "</td>";
                                                cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                                                model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                                            }
                                            else
                                            {
                                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                                table += "<td style='padding: 2px'>&nbsp;</td>";
                                                cpTotalDebit = cpTotalDebit + LcAmount;
                                                model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                                            }
                                            table += "</tr>";
                                            iRowNumber++;
                                        }
                                    }
                                }
                                #endregion

                                #region Agent charges
                                if (Common.toString(SysPrefs.TransactionAgentCommissionAccount) != "" && Common.toString(model.HoldGLReferenceID) != "")
                                {
                                    if (Common.toString(SysPrefs.TransactionAdminCommissionAccount) != Common.toString(SysPrefs.TransactionAgentCommissionAccount))
                                    {
                                        string sqlAgentCharges = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode='" + SysPrefs.TransactionAgentCommissionAccount + "' order by LocalCurrencyAmount Desc";
                                        DataTable objAgentCharges = DBUtils.GetDataTable(sqlAgentCharges);
                                        if (objAgentCharges != null && objAgentCharges.Rows.Count > 0)
                                        {
                                            foreach (DataRow drAgentCharges in objAgentCharges.Rows)
                                            {
                                                //Debit Agent account charges
                                                if (!string.IsNullOrEmpty(drAgentCharges["ForeignCurrencyAmount"].ToString()))
                                                {
                                                    FC = drAgentCharges["ForeignCurrencyAmount"].ToString();
                                                }
                                                FcAmount = Math.Abs(Convert.ToDecimal(FC));
                                                table += "<tr><td style='padding: 2px' valign='top'><strong>" + drAgentCharges["AccountCode"].ToString() + "-" + drAgentCharges["AccountName"].ToString() + "</strong>" + "<br/>" + drAgentCharges["Memo"].ToString() + "<br/>" + drAgentCharges["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAgentCharges["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAgentCharges["GLReferenceID"].ToString() + "</td>";
                                                table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAgentCharges["ForeignCurrencyISOCode"].ToString() + "</td>";
                                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";

                                                if (!string.IsNullOrEmpty(drAgentCharges["LocalCurrencyAmount"].ToString()))
                                                {
                                                    LC = drAgentCharges["LocalCurrencyAmount"].ToString();
                                                }
                                                LcAmount = -1 * Convert.ToDecimal(LC);

                                                if (LcAmount < 0)
                                                {
                                                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "</td>";
                                                    cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                                                    model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                                                }
                                                else
                                                {
                                                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                                    cpTotalDebit = cpTotalDebit + LcAmount;
                                                    model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                                                }
                                                table += "</tr>";
                                                iRowNumber++;

                                                //Credit Agent account charges

                                                if (!string.IsNullOrEmpty(drAgentCharges["ForeignCurrencyAmount"].ToString()))
                                                {
                                                    FC = drAgentCharges["ForeignCurrencyAmount"].ToString();
                                                }
                                                FcAmount = Math.Abs(Convert.ToDecimal(FC));
                                                table += "<tr><td style='padding: 2px' valign='top'><strong>" + myAgentAccountCode + "-" + myAgentAccountName + "</strong>" + "<br/>" + drAgentCharges["Memo"].ToString() + "<br/>" + drAgentCharges["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAgentCharges["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAgentCharges["GLReferenceID"].ToString() + "</td>";
                                                table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAgentCharges["ForeignCurrencyISOCode"].ToString() + "</td>";
                                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";

                                                if (!string.IsNullOrEmpty(drAgentCharges["LocalCurrencyAmount"].ToString()))
                                                {
                                                    LC = drAgentCharges["LocalCurrencyAmount"].ToString();
                                                }
                                                LcAmount = 1 * Convert.ToDecimal(LC);

                                                if (LcAmount < 0)
                                                {
                                                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "</td>";
                                                    cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                                                    model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                                                }
                                                else
                                                {
                                                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                                    cpTotalDebit = cpTotalDebit + LcAmount;
                                                    model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                                                }
                                                table += "</tr>";
                                                iRowNumber++;
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            #endregion



                        }
                        iRowNumber++;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(MyRow["ForeignCurrencyAmount"].ToString()))
                        {
                            FC = MyRow["ForeignCurrencyAmount"].ToString();
                        }
                        FcAmount = Math.Abs(Convert.ToDecimal(FC));
                        table += "<tr><td style='padding: 2px' valign='top'><strong>" + MyRow["AccountCode"].ToString() + "-" + MyRow["AccountName"].ToString() + "</strong>" + "<br/>" + MyRow["Memo"].ToString() + "<br/>" + MyRow["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + MyRow["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + MyRow["GLReferenceID"].ToString() + "</td>";
                        table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + MyRow["ForeignCurrencyISOCode"].ToString() + "</td>";
                        table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";

                        if (!string.IsNullOrEmpty(MyRow["LocalCurrencyAmount"].ToString()))
                        {
                            LC = MyRow["LocalCurrencyAmount"].ToString();
                        }
                        LcAmount = -1 * Convert.ToDecimal(LC);

                        if (LcAmount < 0)
                        {
                            table += "<td style='padding: 2px'>&nbsp;</td>";
                            table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "</td>";
                            cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                            model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                        }
                        else
                        {
                            table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                            table += "<td style='padding: 2px'>&nbsp;</td>";
                            cpTotalDebit = cpTotalDebit + LcAmount;
                            model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                        }
                        table += "</tr>";
                        iRowNumber++;
                    }

                }

            }
            else
            {
                model.Table = Common.GetAlertMessage(1, "Data not found.");
            }

            table += "</tbody>";
            model.Table = table;
            return model;
        }
        public ActionResult Submit(CancelPaymentViewModel model)
        {
            string StrQury = "";
            model.hideContent = true;
            bool isValid = true;
            if (!string.IsNullOrEmpty(model.GeneralAccount))
            {
                if (!string.IsNullOrEmpty(model.Description) && !string.IsNullOrEmpty(model.LCCharges))
                {
                    if (Common.toString(model.CurrencyISOCode) != Common.toString(SysPrefs.DefaultCurrency))
                    {
                        if (!string.IsNullOrEmpty(model.FCCharges))
                        {
                            isValid = true;
                        }
                    }
                    else
                    {
                        isValid = true;
                    }
                }
                else
                {
                    isValid = false;
                }
            }
            if (string.IsNullOrEmpty(model.CustomerTransactionID) && string.IsNullOrEmpty(model.TransactionID))
            {
                isValid = false;
            }
            if (isValid)
            {
                decimal FcAmount = 0.0m;
                decimal LcAmount = 0.0m;
                model.isPosted = false;
                if (!string.IsNullOrEmpty(model.cpGLReferenceID))
                {
                    if (Common.toString(model.Status).Trim() != "Cancelled" && Common.toString(model.CustomerPrefix).Trim() != "")
                    {
                        DataTable dtAdminCharges = new DataTable();
                        DataTable dtAgentCharges = new DataTable();

                        if (model.Status == "Paid")
                        {
                            StrQury = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, GLReferenceID  from GLTransactions where GLReferenceID='" + model.PaidGLReferenceID + "' order by LocalCurrencyAmount Desc";
                        }
                        else
                        {
                            StrQury = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, GLReferenceID  from GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' order by LocalCurrencyAmount Desc";
                        }
                        DataTable objGLTransView = DBUtils.GetDataTable(StrQury);
                        if (objGLTransView != null && objGLTransView.Rows.Count > 0)
                        {
                            Guid guid = Guid.NewGuid();
                            string cpUserRefNo = guid.ToString();

                            #region Reversal of existing voucher

                            //// string myConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
                            using (SqlConnection connection = Common.getConnection())
                            {
                                SqlCommand command = connection.CreateCommand();
                                SqlTransaction transaction = null;
                                try
                                {
                                    // BeginTransaction() Requires Open Connection
                                    //// connection.Open();
                                    transaction = connection.BeginTransaction();
                                    var myPostingDate = SysPrefs.PostingDate;
                                    string myDefaultCurrency = SysPrefs.DefaultCurrency;
                                    foreach (DataRow MyRow in objGLTransView.Rows)
                                    {
                                        command.Parameters.Clear();
                                        string myHoldAccountCode = "";
                                        string sqlListCurrencies = "select HoldAccountCode from ListCurrencies where CurrencyISOCode=@CurrencyISOCode";
                                        command.Transaction = transaction;
                                        command.Parameters.AddWithValue("@CurrencyISOCode", Common.toString(MyRow["ForeignCurrencyISOCode"]).Trim());
                                        command.CommandText = sqlListCurrencies;
                                        if (command.ExecuteScalar() != null)
                                        {
                                            myHoldAccountCode = command.ExecuteScalar().ToString();
                                        }
                                        command.Parameters.Clear();
                                        if (model.Status == "Paid" && myHoldAccountCode == Common.toString(MyRow["AccountCode"]).Trim())
                                        {
                                            //model.CustomerPrefix
                                            string myAgentCurrencyIsoCode = "";
                                            string myAgentAccountCode = "";
                                            string sqlGLChartOfAccounts = "select CurrencyISOCode, AccountCode from  GLChartOfAccounts where prefix=@prefix";
                                            command.Transaction = transaction;
                                            command.Parameters.AddWithValue("@prefix", model.CustomerPrefix);
                                            command.CommandText = sqlGLChartOfAccounts;

                                            DataTable dtCoA = new DataTable();
                                            SqlDataAdapter DataAdaptter = new SqlDataAdapter(command);
                                            DataAdaptter.Fill(dtCoA);
                                            if (dtCoA != null)
                                            {
                                                if (dtCoA.Rows.Count > 0)
                                                {
                                                    myAgentCurrencyIsoCode = Common.toString(dtCoA.Rows[0]["CurrencyISOCode"]);
                                                    myAgentAccountCode = Common.toString(dtCoA.Rows[0]["AccountCode"]);
                                                }
                                            }
                                            command.Parameters.Clear();
                                            string sqlQury = "select * from GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode ='" + myAgentAccountCode + "' and LocalCurrencyAmount='" + MyRow["LocalCurrencyAmount"].ToString() + "'";
                                            DataTable dtAgent = DBUtils.GetDataTable(sqlQury);
                                            if (dtAgent != null && dtAgent.Rows.Count > 0)
                                            {
                                                ///TransactionNumber = dtAgent.Rows[0]["TransactionNumber"].ToString();
                                                string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                command.Transaction = transaction;
                                                command.CommandTimeout = 60;
                                                string accCode = dtAgent.Rows[0]["AccountCode"].ToString();
                                                FcAmount = -1 * Convert.ToDecimal(dtAgent.Rows[0]["ForeignCurrencyAmount"].ToString());
                                                LcAmount = -1 * Convert.ToDecimal(dtAgent.Rows[0]["LocalCurrencyAmount"].ToString());
                                                command.Parameters.AddWithValue("@AccountCode", dtAgent.Rows[0]["AccountCode"].ToString());
                                                command.Parameters.AddWithValue("@GLTransactionTypeiCode", dtAgent.Rows[0]["AccountTypeiCode"].ToString());
                                                command.Parameters.AddWithValue("@Memo", Common.toString(dtAgent.Rows[0]["Memo"]) + " (cancellation)");
                                                command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                command.Parameters.AddWithValue("@ExchangeRate", dtAgent.Rows[0]["ExchangeRate"].ToString());
                                                command.Parameters.AddWithValue("@ForeignCurrencyISOCode", dtAgent.Rows[0]["ForeignCurrencyISOCode"].ToString());
                                                command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                command.CommandText = sqlMyEntry;
                                                command.ExecuteNonQuery();
                                            }
                                            if (model.Status == "Paid")
                                            {
                                                #region Revert Admin & agent charges already paid
                                                if (Common.toString(SysPrefs.TransactionAdminCommissionAccount) != "" && Common.toString(model.HoldGLReferenceID) != "")
                                                {
                                                    string sqlAdminCharges = "select TransactionID,TransactionNumber, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode='" + SysPrefs.TransactionAdminCommissionAccount + "' order by LocalCurrencyAmount Desc";
                                                    dtAdminCharges = DBUtils.GetDataTable(sqlAdminCharges);
                                                    if (dtAdminCharges != null)
                                                    {
                                                        if (dtAdminCharges.Rows.Count > 0)
                                                        {
                                                            foreach (DataRow drAdminCharges in dtAdminCharges.Rows)
                                                            {
                                                                command.Parameters.Clear();
                                                                string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                                FcAmount = -1 * Convert.ToDecimal(drAdminCharges["ForeignCurrencyAmount"].ToString());
                                                                LcAmount = -1 * Convert.ToDecimal(drAdminCharges["LocalCurrencyAmount"].ToString());
                                                                command.Parameters.AddWithValue("@AccountCode", drAdminCharges["AccountCode"].ToString());
                                                                command.Parameters.AddWithValue("@GLTransactionTypeiCode", drAdminCharges["AccountTypeiCode"].ToString());
                                                                command.Parameters.AddWithValue("@Memo", Common.toString(drAdminCharges["Memo"]) + " (cancellation)");
                                                                command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                                command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                                command.Parameters.AddWithValue("@ExchangeRate", drAdminCharges["ExchangeRate"].ToString());
                                                                command.Parameters.AddWithValue("@ForeignCurrencyISOCode", drAdminCharges["ForeignCurrencyISOCode"].ToString());
                                                                command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                                command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                                command.CommandText = sqlMyEntry;
                                                                command.ExecuteNonQuery();
                                                                command.Parameters.Clear();

                                                                sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                                FcAmount = 1 * Convert.ToDecimal(drAdminCharges["ForeignCurrencyAmount"].ToString());
                                                                LcAmount = 1 * Convert.ToDecimal(drAdminCharges["LocalCurrencyAmount"].ToString());
                                                                command.Parameters.AddWithValue("@AccountCode", myAgentAccountCode);
                                                                command.Parameters.AddWithValue("@GLTransactionTypeiCode", drAdminCharges["AccountTypeiCode"].ToString());
                                                                command.Parameters.AddWithValue("@Memo", Common.toString(drAdminCharges["Memo"]) + " (cancellation)");
                                                                command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                                command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                                command.Parameters.AddWithValue("@ExchangeRate", drAdminCharges["ExchangeRate"].ToString());
                                                                command.Parameters.AddWithValue("@ForeignCurrencyISOCode", drAdminCharges["ForeignCurrencyISOCode"].ToString());
                                                                command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                                command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                                command.CommandText = sqlMyEntry;
                                                                command.ExecuteNonQuery();
                                                                command.Parameters.Clear();
                                                            }
                                                        }
                                                    }
                                                }
                                                if (Common.toString(SysPrefs.TransactionAgentCommissionAccount) != "" && Common.toString(model.HoldGLReferenceID) != "")
                                                {
                                                    if (Common.toString(SysPrefs.TransactionAdminCommissionAccount) != Common.toString(SysPrefs.TransactionAgentCommissionAccount))
                                                    {
                                                        string sqlAgentCharges = "select TransactionID,TransactionNumber, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode='" + SysPrefs.TransactionAgentCommissionAccount + "' order by LocalCurrencyAmount Desc";
                                                        dtAgentCharges = DBUtils.GetDataTable(sqlAgentCharges);
                                                        if (dtAgentCharges != null)
                                                        {

                                                            if (dtAgentCharges.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow drAgentCharges in dtAgentCharges.Rows)
                                                                {
                                                                    command.Parameters.Clear();
                                                                    string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                                    FcAmount = -1 * Convert.ToDecimal(drAgentCharges["ForeignCurrencyAmount"].ToString());
                                                                    LcAmount = -1 * Convert.ToDecimal(drAgentCharges["LocalCurrencyAmount"].ToString());
                                                                    command.Parameters.AddWithValue("@AccountCode", drAgentCharges["AccountCode"].ToString());
                                                                    command.Parameters.AddWithValue("@GLTransactionTypeiCode", drAgentCharges["AccountTypeiCode"].ToString());
                                                                    command.Parameters.AddWithValue("@Memo", Common.toString(drAgentCharges["Memo"]) + " (cancellation)");
                                                                    command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                                    command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                                    command.Parameters.AddWithValue("@ExchangeRate", drAgentCharges["ExchangeRate"].ToString());
                                                                    command.Parameters.AddWithValue("@ForeignCurrencyISOCode", drAgentCharges["ForeignCurrencyISOCode"].ToString());
                                                                    command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                                    command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                                    command.CommandText = sqlMyEntry;
                                                                    command.ExecuteNonQuery();
                                                                    command.Parameters.Clear();

                                                                    sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                                    FcAmount = 1 * Convert.ToDecimal(drAgentCharges["ForeignCurrencyAmount"].ToString());
                                                                    LcAmount = 1 * Convert.ToDecimal(drAgentCharges["LocalCurrencyAmount"].ToString());
                                                                    command.Parameters.AddWithValue("@AccountCode", myAgentAccountCode);
                                                                    command.Parameters.AddWithValue("@GLTransactionTypeiCode", drAgentCharges["AccountTypeiCode"].ToString());
                                                                    command.Parameters.AddWithValue("@Memo", Common.toString(drAgentCharges["Memo"]) + " (cancellation)");
                                                                    command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                                    command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                                    command.Parameters.AddWithValue("@ExchangeRate", drAgentCharges["ExchangeRate"].ToString());
                                                                    command.Parameters.AddWithValue("@ForeignCurrencyISOCode", drAgentCharges["ForeignCurrencyISOCode"].ToString());
                                                                    command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                                    command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                                    command.CommandText = sqlMyEntry;
                                                                    command.ExecuteNonQuery();
                                                                    command.Parameters.Clear();
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion
                                            }
                                        }
                                        else
                                        {
                                            string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                            command.Transaction = transaction;
                                            command.CommandTimeout = 60;
                                            string accCode = MyRow["AccountCode"].ToString();
                                            FcAmount = -1 * Convert.ToDecimal(MyRow["ForeignCurrencyAmount"].ToString());
                                            LcAmount = -1 * Convert.ToDecimal(MyRow["LocalCurrencyAmount"].ToString());
                                            command.Parameters.AddWithValue("@AccountCode", MyRow["AccountCode"].ToString());
                                            command.Parameters.AddWithValue("@GLTransactionTypeiCode", MyRow["AccountTypeiCode"].ToString());
                                            command.Parameters.AddWithValue("@Memo", Common.toString(MyRow["Memo"]) + " (cancellation)");
                                            command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                            command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                            command.Parameters.AddWithValue("@ExchangeRate", MyRow["ExchangeRate"].ToString());
                                            command.Parameters.AddWithValue("@ForeignCurrencyISOCode", MyRow["ForeignCurrencyISOCode"].ToString());
                                            command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                            command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                            command.CommandText = sqlMyEntry;
                                            command.ExecuteNonQuery();
                                        }
                                        command.Parameters.Clear();
                                    }

                                    decimal myAdminCharges = Common.toDecimal(model.LCCharges);
                                    //*** Cancellation Charges Transactions ***//

                                    #region Cancellation Charges Transactions
                                    if (myAdminCharges > 0)
                                    {
                                        string myChargesAccountCode = model.GeneralAccount;
                                        if (myChargesAccountCode != "")
                                        {
                                            #region credit charges account 
                                            string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                            decimal ExchangeRate = 1;
                                            string ForeignCurrencyISOCode = "";
                                            decimal ForeignCurrencyAmount = 0;

                                            string sqlGLChartOfAccounts = "select CurrencyISOCode from  GLChartOfAccounts where AccountCode=@AccountCode";
                                            command.Transaction = transaction;
                                            command.Parameters.AddWithValue("@AccountCode", myChargesAccountCode);
                                            command.CommandText = sqlGLChartOfAccounts;
                                            if (command.ExecuteScalar() != null)
                                            {
                                                ForeignCurrencyISOCode = Common.toString(command.ExecuteScalar());
                                            }

                                            sqlGLChartOfAccounts = "select top 1 ExchangeRate from ExchangeRates where CurrencyISOCode=@CurrencyISOCode and AddedDate =@AddedDate";
                                            command.Transaction = transaction;
                                            command.Parameters.AddWithValue("@CurrencyISOCode", ForeignCurrencyISOCode);
                                            command.Parameters.AddWithValue("@AddedDate", Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00"));
                                            command.CommandText = sqlGLChartOfAccounts;
                                            if (command.ExecuteScalar() != null)
                                            {
                                                ExchangeRate = Common.toDecimal(command.ExecuteScalar());
                                            }
                                            command.Parameters.Clear();
                                            if (ForeignCurrencyISOCode != "" && ExchangeRate > 0)
                                            {
                                                if (!string.IsNullOrEmpty(model.FCCharges))
                                                {
                                                    ForeignCurrencyAmount = Common.toDecimal(model.FCCharges);
                                                }
                                                else
                                                {
                                                    ForeignCurrencyAmount = Common.toDecimal(ExchangeRate * myAdminCharges);
                                                }

                                                command.Parameters.Clear();
                                                ForeignCurrencyAmount = -1 * ForeignCurrencyAmount;
                                                command.Parameters.AddWithValue("@AccountCode", myChargesAccountCode);
                                                command.Parameters.AddWithValue("@GLTransactionTypeiCode", TypeiCode);
                                                command.Parameters.AddWithValue("@Memo", model.Description);
                                                command.Parameters.AddWithValue("@ForeignCurrencyAmount", ForeignCurrencyAmount);
                                                command.Parameters.AddWithValue("@LocalCurrencyAmount", -1 * myAdminCharges);
                                                command.Parameters.AddWithValue("@ExchangeRate", ExchangeRate);
                                                command.Parameters.AddWithValue("@ForeignCurrencyISOCode", ForeignCurrencyISOCode);
                                                command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                command.CommandText = sqlMyEntry;
                                                command.ExecuteNonQuery();
                                                command.Parameters.Clear();
                                            }
                                            else
                                            {
                                                transaction.Rollback();
                                                isValid = false;
                                                model.ErrMessage = "Account currency not setup.";
                                                model.isPosted = false;
                                            }

                                            #endregion

                                            #region Agent debit account 
                                            if (isValid)
                                            {
                                                sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                ExchangeRate = 1;
                                                ForeignCurrencyISOCode = "";
                                                ForeignCurrencyAmount = 0;
                                                string myAgentAccountCode = "";
                                                sqlGLChartOfAccounts = "select CurrencyISOCode, AccountCode from  GLChartOfAccounts where prefix=@prefix";
                                                command.Transaction = transaction;
                                                command.Parameters.AddWithValue("@prefix", model.CustomerPrefix);
                                                command.CommandText = sqlGLChartOfAccounts;

                                                DataTable dtCoA = new DataTable();
                                                SqlDataAdapter DataAdaptter = new SqlDataAdapter(command);
                                                DataAdaptter.Fill(dtCoA);
                                                if (dtCoA != null)
                                                {
                                                    if (dtCoA.Rows.Count > 0)
                                                    {
                                                        ForeignCurrencyISOCode = Common.toString(dtCoA.Rows[0]["CurrencyISOCode"]);
                                                        myAgentAccountCode = Common.toString(dtCoA.Rows[0]["AccountCode"]);
                                                    }
                                                }

                                                sqlGLChartOfAccounts = "select top 1 ExchangeRate from ExchangeRates where CurrencyISOCode=@CurrencyISOCode and AddedDate =@AddedDate";
                                                command.Transaction = transaction;
                                                command.Parameters.AddWithValue("@CurrencyISOCode", ForeignCurrencyISOCode);
                                                command.Parameters.AddWithValue("@AddedDate", Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00"));
                                                command.CommandText = sqlGLChartOfAccounts;
                                                if (command.ExecuteScalar() != null)
                                                {
                                                    ExchangeRate = Common.toDecimal(command.ExecuteScalar());
                                                }
                                                command.Parameters.Clear();
                                                if (ForeignCurrencyISOCode != "" && ExchangeRate > 0 && myAgentAccountCode != "")
                                                {
                                                    if (!string.IsNullOrEmpty(model.FCCharges))
                                                    {
                                                        ForeignCurrencyAmount = Common.toDecimal(model.FCCharges);
                                                    }
                                                    else
                                                    {
                                                        ForeignCurrencyAmount = Common.toDecimal(ExchangeRate * myAdminCharges);
                                                    }
                                                    command.Parameters.AddWithValue("@AccountCode", myAgentAccountCode);
                                                    command.Parameters.AddWithValue("@GLTransactionTypeiCode", TypeiCode);
                                                    command.Parameters.AddWithValue("@Memo", model.Description);
                                                    command.Parameters.AddWithValue("@ForeignCurrencyAmount", ForeignCurrencyAmount);
                                                    command.Parameters.AddWithValue("@LocalCurrencyAmount", myAdminCharges);
                                                    command.Parameters.AddWithValue("@ExchangeRate", ExchangeRate);
                                                    command.Parameters.AddWithValue("@ForeignCurrencyISOCode", ForeignCurrencyISOCode);
                                                    command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                    command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                    command.CommandText = sqlMyEntry;
                                                    command.ExecuteNonQuery();
                                                    command.Parameters.Clear();
                                                }
                                                else
                                                {
                                                    transaction.Rollback();
                                                    isValid = false;
                                                    model.ErrMessage = "Agent account or currency not setup.";
                                                    model.isPosted = false;
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                    #endregion

                                    if (isValid)
                                    {
                                        transaction.Commit();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();
                                    model.ErrMessage = Common.GetAlertMessage(1, "Error: " + ex.Message);
                                    model.isPosted = false;
                                }
                                finally
                                {
                                    connection.Close();
                                }
                            }
                            #endregion

                            #region Post for GL
                            if (isValid)
                            {
                                DataTable lstGLEntriesTemp = DBUtils.GetDataTable("select * from GLEntriesTemp where UserRefNo='" + cpUserRefNo + "'");
                                if (lstGLEntriesTemp != null)
                                {
                                    if (lstGLEntriesTemp.Rows.Count > 0)
                                    {
                                        PostingViewModel PostingModel = PostingUtils.SaveGeneralJournelGLEntries(cpUserRefNo, "Transaction cancellation", "22CANCELENTRY", lstGLEntriesTemp, Profile.Id.ToString());
                                        if (PostingModel.isPosted)
                                        {
                                            DBUtils.ExecuteSQL("Update CustomerTransactions set CancelledGLReferenceID = '" + PostingModel.GLReferenceID + "', CancelledDate=GETDATE() where TransactionID = '" + model.TransactionID + "'");
                                            model.isPosted = false;
                                            model.hideContent = false;
                                            model.ErrMessage = PostingModel.ErrMessage;
                                        }
                                        else
                                        {
                                            model.ErrMessage = PostingModel.ErrMessage;
                                        }
                                    }
                                    else
                                    {
                                        model.ErrMessage = Common.GetAlertMessage(1, "No temp data found.. cp:" + cpUserRefNo);
                                    }
                                }
                                else
                                {
                                    model.ErrMessage = Common.GetAlertMessage(1, "No temp data found. cp:" + cpUserRefNo);
                                }
                            }
                            else
                            {
                                model.ErrMessage = Common.GetAlertMessage(1, "No valid data");
                            }
                            #endregion
                        }
                        else
                        {
                            model.ErrMessage = Common.GetAlertMessage(1, "Data not found. id:" + model.cpGLReferenceID);
                        }
                    }
                    else
                    {
                        model.CancellationMessage = Common.GetAlertMessage(1, "Transaction already cancelled.");
                    }
                }
                else
                {
                    model.ErrMessage = Common.GetAlertMessage(1, "Invalid Reference ID");
                }
            }
            else
            {
                model.ErrMessage = Common.GetAlertMessage(1, "Incomplete data. please provide charges description and amount");
            }
            return View("~/Views/Postings/CancelPayment.cshtml", model);
        }
        #endregion

        #region GeneralJournal Delete
        public ActionResult GeneralJournalDelete()
        {
            PostingViewModel postingViewModel = new PostingViewModel();
            postingViewModel.isListPosted = true;
            postingViewModel.isPosted = true;
            postingViewModel.isBtnShow = false;
            postingViewModel.FromDate = Common.toDateTime(SysPrefs.PostingDate).AddDays(-30);
            postingViewModel.ToDate = Common.toDateTime(SysPrefs.PostingDate);
            return View("~/Views/Postings/GeneralJournalDelete.cshtml", postingViewModel);
        }
        [HttpPost]
        public ActionResult GeneralJournalDelete(string VoucherNumber)
        {
            PostingViewModel postingViewModel = new PostingViewModel();
            if (!string.IsNullOrEmpty(VoucherNumber))
            {
                string mySql = "Select GLReferenceID from GLReferences";
                mySql += " where VoucherNumber ='" + VoucherNumber + "'";
                string id = Common.toString(DBUtils.executeSqlGetSingle(mySql));
                //TransViewModel pGLTransViewModel = new TransViewModel();
                if (!string.IsNullOrEmpty(id))
                {
                    postingViewModel.cpRefID = id;
                    postingViewModel.HtmlTable1 = BuildDynamicTable1(Common.toInt(postingViewModel.cpRefID));
                    postingViewModel.isListPosted = false;
                    string Records = "Select VoucherNumber, AuthorizedBy, AuthorizedDate, DatePosted, PostedBy from GLReferences where GLReferenceID=" + postingViewModel.cpRefID + "";
                    DataTable transTable = DBUtils.GetDataTable(Records);
                    if (transTable != null && transTable.Rows.Count > 0)
                    {
                        foreach (DataRow RecordTable in transTable.Rows)
                        {
                            postingViewModel.AuthorizedBy = Common.GetUserFullName(Common.toString(RecordTable["AuthorizedBy"]));
                            postingViewModel.AuthorizedDate = Common.toDateTime(Common.toDateTime(RecordTable["AuthorizedDate"]).ToShortDateString());
                            postingViewModel.DatePosted = Common.toDateTime(Common.toDateTime(RecordTable["DatePosted"]).ToShortDateString());
                            postingViewModel.PostedBy = Common.GetUserFullName(Common.toString(RecordTable["PostedBy"]));
                            postingViewModel.VoucherNumber = Common.toString(RecordTable["VoucherNumber"]);
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
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Please provide voucher no.");
            }
            return View("~/Views/Postings/GeneralJournalDelete.cshtml", postingViewModel);
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
                        table += "<td valign='top' align='right'>" + Math.Round(FcAmount, 2) + "</td>";
                        string LC = "0.00";
                        if (!string.IsNullOrEmpty(MyRow["LocalCurrencyAmount"].ToString()))
                        {
                            LC = MyRow["LocalCurrencyAmount"].ToString();
                        }
                        decimal LcAmount = Convert.ToDecimal(LC);
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
            table += "<td colspan='2'></td>";
            table += "<td style='background-color:#1c3a70; color:#FFF;'>Total:</td>";
            table += "<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + DisplayUtils.GetSystemAmountFormat(cpTotalDebit) + "</td>";
            table += "<td style='background-color:#1c3a70; color:#FFF; text-align:right;'>" + DisplayUtils.GetSystemAmountFormat(cpTotalCredit) + "</td>";
            table += "</tr>";
            table += "</tfoot>";
            table += "</table>";
            return table;
        }
        public ActionResult DeleteVoucher(PostingViewModel pModel)
        {
            using (SqlConnection connection = Common.getConnection())
            {
                SqlTransaction transaction = null;
                try
                {
                    transaction = connection.BeginTransaction();
                    SqlCommand command = connection.CreateCommand();
                    ApplicationUser Profile = ApplicationUser.GetUserProfile();
                    if (Common.ValidatePinCode(Profile.Id, pModel.PinCode))
                    {
                        bool isAllowed = false;
                        string PostedBy = DBUtils.executeSqlGetSingle("Select PostedBy from GLReferences where [glreferenceid]=" + pModel.cpRefID + "");
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
                                string myRefferenceNo = Common.toString(pModel.cpRefID);
                                if (myRefferenceNo != "")
                                {
                                    string AccountTypeiCode = "";
                                    string TransactionNumber = "";
                                    string PaidGLReferenceID = "";

                                    string sqlGLTransactions = "Select top 1 AccountTypeiCode, TransactionNumber from GLTransactions where GLReferenceID=@GLReferenceID";
                                    command.Transaction = transaction;
                                    command.Parameters.AddWithValue("@GLReferenceID", myRefferenceNo);
                                    command.CommandText = sqlGLTransactions;
                                    DataTable dtGLTransactions = new DataTable();
                                    SqlDataAdapter DataAdaptter = new SqlDataAdapter(command);
                                    DataAdaptter.Fill(dtGLTransactions);
                                    if (dtGLTransactions != null)
                                    {
                                        if (dtGLTransactions.Rows.Count > 0)
                                        {
                                            AccountTypeiCode = Common.toString(dtGLTransactions.Rows[0]["AccountTypeiCode"]);
                                            TransactionNumber = Common.toString(dtGLTransactions.Rows[0]["TransactionNumber"]);
                                        }
                                        DataAdaptter.Dispose();
                                    }
                                    command.Parameters.Clear();

                                    string updateCustomerTrans = "";
                                    if (AccountTypeiCode != "" && TransactionNumber != "")
                                    {
                                        if (AccountTypeiCode == "22CANCELENTRY")
                                        {
                                            string sqlCustomerTransactions = "Select top 1 PaidGLReferenceID from CustomerTransactions where paymentno=@TransactionNumber";
                                            command.Transaction = transaction;
                                            command.Parameters.AddWithValue("@TransactionNumber", TransactionNumber);
                                            command.CommandText = sqlCustomerTransactions;
                                            DataTable dtCustomerTransactions = new DataTable();
                                            SqlDataAdapter daCustomerTransactions = new SqlDataAdapter(command);
                                            daCustomerTransactions.Fill(dtCustomerTransactions);
                                            if (dtCustomerTransactions != null)
                                            {
                                                if (dtCustomerTransactions.Rows.Count > 0)
                                                {
                                                    PaidGLReferenceID = Common.toString(dtCustomerTransactions.Rows[0]["PaidGLReferenceID"]);
                                                }
                                                daCustomerTransactions.Dispose();
                                            }
                                            command.Parameters.Clear();

                                            command.Transaction = transaction;
                                            //command.CommandTimeout = 60;
                                            updateCustomerTrans = "update CustomerTransactions set CancelledGLReferenceID = null ";
                                            if (Common.toInt(PaidGLReferenceID) > 0)
                                            {
                                                updateCustomerTrans += ", status = 'Paid' ";
                                            }
                                            else
                                            {
                                                updateCustomerTrans += ", status = 'Unpaid' ";
                                            }
                                            updateCustomerTrans += " where paymentno = '" + TransactionNumber + "'";
                                            command.CommandText = updateCustomerTrans;
                                            command.ExecuteNonQuery();
                                            command.Parameters.Clear();
                                        }
                                    }

                                    command.Transaction = transaction;
                                    //command.CommandTimeout = 60;
                                    updateCustomerTrans = "update CustomerTransactions set PaidGLReferenceID = null, status='Unpaid' where PaidGLReferenceID = '" + myRefferenceNo + "'";
                                    command.CommandText = updateCustomerTrans;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();

                                    command.Transaction = transaction;
                                    string deleteCustomerTrans = "delete from CustomerTransactions where HoldGLReferenceID = '" + myRefferenceNo + "'";
                                    command.CommandText = deleteCustomerTrans;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();

                                    command.Transaction = transaction;
                                    string SqlDeletetransaction = "Delete from GLTransactions where GLReferenceID=@GLReferenceID";
                                    command.Parameters.AddWithValue("@GLReferenceID", pModel.cpRefID);
                                    command.CommandText = SqlDeletetransaction;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();

                                    command.Transaction = transaction;
                                    string SqlDeleteVoucher = "Delete From GLReferences where GLReferenceID=@GLReferenceID;";
                                    command.Parameters.AddWithValue("@GLReferenceID", myRefferenceNo);
                                    command.CommandText = SqlDeleteVoucher;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();

                                }
                                command.Transaction = transaction;
                                string updateRef = "Update GLReferences Set VoucherNumber=VoucherNumber-1 Where GLReferenceID >'" + pModel.cpRefID + "'";
                                command.CommandText = updateRef;
                                command.ExecuteNonQuery();
                                command.Parameters.Clear();

                                pModel.isPosted = true;
                                pModel.ErrMessage = Common.GetAlertMessage(0, "&nbsp &nbsp &nbsp Voucher deleted succcessfully.");
                                transaction.Commit();
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                pModel.ErrMessage = Common.GetAlertMessage(1, "Error occured.");
                            }
                        }
                        else
                        {
                            transaction.Rollback();
                            pModel.ErrMessage = Common.GetAlertMessage(1, "Sorry! same user can not deleted voucher posted by same user.");
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                        pModel.isPosted = true;
                        pModel.ErrMessage = Common.GetAlertMessage(1, "Invalid pin code provided.");
                    }
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                }
                finally
                {
                    connection.Close();
                }
            }
            return View("~/Views/Postings/GeneralJournalDelete.cshtml", pModel);
        }
        #endregion

        #region Delete Voucher Range
        public JsonResult GetTypeiCode([DataSourceRequest] DataSourceRequest request)
        {
            var TypeiCode = PostingUtils.GetTypeiCodes();
            return Json(TypeiCode, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GeneralJournalDeleteRange(PostingViewModel postingViewModel)
        {
            string table = "";
            try
            {
                if (Common.toDateTime(postingViewModel.FromDate) != null && Common.toDateTime(postingViewModel.ToDate) != null && !string.IsNullOrEmpty(Common.toString(postingViewModel.TypeiCode)))
                {
                    string mySql = "Select VoucherNumber, ReferenceNo, DatePosted, GLReferenceID,(Select Description From ConfigItemsData Where ConfigItemsData.ItemsDataCode = GLReferences.TypeiCode) As TypeiCode from GLReferences ";
                    mySql += " where GLReferenceID in (select  GLReferenceID from GLReferences where GLReferences.DatePosted  >= '" + Common.toDateTime(postingViewModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and GLReferences.DatePosted  <= '" + Common.toDateTime(postingViewModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and GLReferences.TypeiCode='" + postingViewModel.TypeiCode + "') order by GLReferenceID DESC";
                    DataTable dtRange = DBUtils.GetDataTable(mySql);
                    if (dtRange != null && dtRange.Rows.Count > 0)
                    {
                        table += "<table id='tblMain' border='1' width='100%' cellpadding='5'>";
                        table += "<thead>";
                        table += " <tr>";
                        //table += "<th style='background-color:#007bff; color:#FFF;'><input id='chk' type='checkbox' value=" + 1 + " onclick='checkAll(this);'></th>";
                        table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Date</th>";
                        table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Voucher #</th>";
                        table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Reference ID</th>";
                        table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Type</th>";
                        table += "</tr>";
                        table += "<thead>";
                        table += "<tbody>";
                        foreach (DataRow dr in dtRange.Rows)
                        {
                            table += "<tr>";
                            //table += "<tr id='" + Common.toString(Item["GLReferenceID"]) + "'><td><input id='chk_" + Common.toString(Item["GLReferenceID"]) + "' type='checkbox' Class='case' value=" + Common.toString(Item["GLReferenceID"]) + " onclick='checkClick(this);'></td>";
                            table += "<td>" + dr["DatePosted"].ToString().Substring(0, 10) + "</td>";
                            table += "<td>" + dr["VoucherNumber"].ToString() + "</td>";
                            table += "<td>" + dr["GLReferenceID"].ToString() + "</td>";
                            table += "<td>" + dr["TypeiCode"].ToString() + "</td></tr>";
                        }
                        postingViewModel.isPosted = false;
                        postingViewModel.isListPosted = true;
                        postingViewModel.isBtnShow = true;
                    }
                    else
                    {
                        postingViewModel.isPosted = true;
                        postingViewModel.ErrMessage = Common.GetAlertMessage(1, "No record found.");
                    }
                }
                else
                {
                    postingViewModel.isPosted = true;
                    postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Please select voucher type.");
                }
                table += "</tbody>";
                table += "</table>";
                postingViewModel.ListTable = table;
            }
            catch (Exception ex)
            {
                postingViewModel.isPosted = true;
                postingViewModel.ErrMessage = "An error occured: " + ex.ToString();
            }
            return View("~/Views/Postings/GeneralJournalDelete.cshtml", postingViewModel);
        }

        [HttpPost]
        public ActionResult DeleteVoucherRange(PostingViewModel pModel)
        {
            ApplicationUser Profile = ApplicationUser.GetUserProfile();
            bool isDeleted = false;
            if (Common.ValidatePinCode(Profile.Id, pModel.PinCode))
            {
                using (SqlConnection connection = Common.getConnection())
                {
                    SqlTransaction transaction = null;
                    try
                    {
                        transaction = connection.BeginTransaction();
                        SqlCommand command = connection.CreateCommand();
                        if (Common.toDateTime(pModel.FromDate) != null && Common.toDateTime(pModel.ToDate) != null && !string.IsNullOrEmpty(Common.toString(pModel.TypeiCode)))
                        {
                            string FirstGLReferenceID = "";
                            string LastGLReferenceID = "";
                            int FirstVoucherNumber = 0;
                            string sqlFirst = "Select top 1 GLReferenceID , VoucherNumber from GLReferences where GLReferences.DatePosted  >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and GLReferences.DatePosted  <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and GLReferences.TypeiCode='" + pModel.TypeiCode + "' order by GLReferenceID";
                            DataTable dtFirst = DBUtils.GetDataTable(sqlFirst);
                            if (dtFirst != null)
                            {
                                if (dtFirst.Rows.Count > 0)
                                {
                                    FirstGLReferenceID = Common.toString(dtFirst.Rows[0]["GLReferenceID"]);
                                    FirstVoucherNumber = Common.toInt(dtFirst.Rows[0]["VoucherNumber"]);
                                }
                            }
                            string sqlLast = "Select top 1 GLReferenceID from GLReferences where GLReferences.DatePosted  >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and GLReferences.DatePosted  <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and GLReferences.TypeiCode='" + pModel.TypeiCode + "' order by GLReferenceID Desc";
                            LastGLReferenceID = DBUtils.executeSqlGetSingle(sqlLast);

                            //string sqlNextRefID = "Select top 1 GLReferenceID ,VoucherNumber from GLReferences where GLReferenceID > " + LastGLReferenceID + " order by GLReferenceID";
                            //DataTable dtNextRefID = DBUtils.GetDataTable(sqlNextRefID);
                            //if (dtNextRefID != null)
                            //{
                            //    if (dtNextRefID.Rows.Count > 0)
                            //    {
                            //        NextRefID = Common.toString(dtNextRefID.Rows[0]["GLReferenceID"]);
                            //        NextVoucherNumber = Common.toInt(dtNextRefID.Rows[0]["VoucherNumber"]);
                            //    }
                            //}

                            if (FirstVoucherNumber > 0 && Common.toString(FirstGLReferenceID).Trim() != "" && Common.toString(LastGLReferenceID).Trim() != "")
                            {
                                try
                                {
                                    string updateCustomerTrans = "";
                                    if (pModel.TypeiCode == "22CANCELENTRY")
                                    {
                                        command.Transaction = transaction;
                                        updateCustomerTrans = "update CustomerTransactions set CancelledGLReferenceID = null, status='Unpaid' where PaidGLReferenceID is null and CancelledGLReferenceID in (select  GLReferenceID from GLReferences where  DatePosted >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and DatePosted <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and TypeiCode='" + pModel.TypeiCode + "')";
                                        command.CommandText = updateCustomerTrans;
                                        command.ExecuteNonQuery();
                                        command.Parameters.Clear();

                                        command.Transaction = transaction;
                                        updateCustomerTrans = "update CustomerTransactions set CancelledGLReferenceID = null, status='Paid' where PaidGLReferenceID is not null and CancelledGLReferenceID in (select  GLReferenceID from GLReferences where  DatePosted >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and DatePosted <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and TypeiCode='" + pModel.TypeiCode + "')";
                                        command.CommandText = updateCustomerTrans;
                                        command.ExecuteNonQuery();
                                        command.Parameters.Clear();
                                    }

                                    command.Transaction = transaction;
                                    //command.CommandTimeout = 60;
                                    updateCustomerTrans = "update CustomerTransactions set PaidGLReferenceID = null, status='Unpaid'  where PaidGLReferenceID in (select  GLReferenceID from GLReferences where  DatePosted >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and DatePosted <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and TypeiCode='" + pModel.TypeiCode + "')";
                                    command.CommandText = updateCustomerTrans;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();

                                    command.Transaction = transaction;
                                    string deleteCustomerTrans = "delete from CustomerTransactions where HoldGLReferenceID in (select  GLReferenceID from GLReferences where  DatePosted >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and DatePosted <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and TypeiCode='" + pModel.TypeiCode + "')";
                                    command.CommandText = deleteCustomerTrans;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();

                                    command.Transaction = transaction;
                                    string deleteGLTrans = "delete from GLTransactions where GLReferenceID in (select  GLReferenceID  from GLReferences where DatePosted >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and DatePosted <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and TypeiCode='" + pModel.TypeiCode + "')";
                                    command.CommandText = deleteGLTrans;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();

                                    command.Transaction = transaction;
                                    string deleteGLReferences = "delete from GLReferences where  GLReferenceID in (select  GLReferenceID from GLReferences where  DatePosted >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and DatePosted <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and TypeiCode='" + pModel.TypeiCode + "')";
                                    command.CommandText = deleteGLReferences;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();

                                    FirstVoucherNumber = FirstVoucherNumber - 1;
                                    command.Transaction = transaction;
                                    ////string updateGLReferences = " ;WITH cteRows As ( SELECT  GLReferenceID, row_number() over (order by VoucherNumber) as NewOrder FROM GLReferences where GLReferenceID >= '" + FirstGLReferenceID + "' and GLReferenceID <= '" + LastGLReferenceID + "') UPDATE GLReferences SET VoucherNumber = " + FirstVoucherNumber + " + NewOrder FROM GLReferences INNER JOIN cteRows ON cteRows.GLReferenceID = GLReferences.GLReferenceID";
                                    string updateGLReferences = " ;WITH cteRows As ( SELECT  GLReferenceID, row_number() over (order by VoucherNumber) as NewOrder FROM GLReferences where GLReferenceID >= " + FirstGLReferenceID + " ) UPDATE GLReferences SET VoucherNumber = " + FirstVoucherNumber + " + NewOrder FROM GLReferences INNER JOIN cteRows ON cteRows.GLReferenceID = GLReferences.GLReferenceID";
                                    command.CommandText = updateGLReferences;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();
                                    pModel.isPosted = true;
                                    isDeleted = true;
                                    transaction.Commit();
                                    //string sqlPreviousVno = "Select top 1 GLReferenceID,VoucherNumber from GLReferences where GLReferenceID < '" + LastGLReferenceID + "' order by GLReferenceID desc";
                                    //DataTable dtPreviousVno = DBUtils.GetDataTable(sqlPreviousVno);
                                    //if (dtPreviousVno != null)
                                    //{
                                    //    if (dtPreviousVno.Rows.Count > 0)
                                    //    {
                                    //        PreRefID = Common.toString(dtPreviousVno.Rows[0]["GLReferenceID"]);
                                    //        PreVoucherNumber = Common.toInt(dtPreviousVno.Rows[0]["VoucherNumber"]);
                                    //    }
                                    //}
                                    //int Diffrence = NextVoucherNumber - PreVoucherNumber;
                                    //Diffrence = Diffrence - 1;
                                    //transaction = connection.BeginTransaction();
                                    //command.Transaction = transaction;
                                    //string updateRef = "Update GLReferences Set VoucherNumber=VoucherNumber-" + Diffrence + " Where GLReferenceID >'" + LastGLReferenceID + "'";
                                    //command.CommandText = updateRef;
                                    //command.ExecuteNonQuery();
                                    //command.Parameters.Clear();
                                    //transaction.Commit();

                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();
                                    pModel.isPosted = true;
                                    isDeleted = false;
                                    pModel.ErrMessage = Common.GetAlertMessage(1, "Error occured.");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        isDeleted = false;
                        //transaction.Rollback();
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }
            else
            {
                isDeleted = false;
                pModel.isPosted = true;
                pModel.ErrMessage = Common.GetAlertMessage(1, "Invalid pin code provided.");
            }
            if (isDeleted)
            {
                pModel.ErrMessage = Common.GetAlertMessage(0, "&nbsp &nbsp &nbsp Selected Vouchers deleted succcessfully.");
            }
            return View("~/Views/Postings/GeneralJournalDelete.cshtml", pModel);
        }

        //public ActionResult DeleteVoucherRange(PostingViewModel pModel)
        //{
        //    ApplicationUser Profile = ApplicationUser.GetUserProfile();
        //    bool isDeleted = false;
        //    if (Common.ValidatePinCode(Profile.Id, pModel.PinCode))
        //    {
        //        using (SqlConnection connection = Common.getConnection())
        //        {
        //            SqlTransaction transaction = null;
        //            try
        //            {
        //                transaction = connection.BeginTransaction();
        //                SqlCommand command = connection.CreateCommand();
        //                if (!string.IsNullOrEmpty(pModel.StartingRange) && !string.IsNullOrEmpty(pModel.EndingRange))
        //                {
        //                    bool isAllowed = false;
        //                    string mySql = "Select PostedBy,TypeiCode, GLReferenceID from GLReferences";
        //                    mySql += " where GLReferenceID in (select  GLReferenceID from GLReferences where DatePosted >= '" + Common.toDateTime(pModel.FromDate).ToString("yyyy/MM/dd 00:00:00") + "' and DatePosted <= '" + Common.toDateTime(pModel.ToDate).ToString("yyyy/MM/dd 00:00:00") + "' and PostedBy ='" + Profile.Id + "' and TypeiCode='" + pModel.TypeiCode + "') order by GLReferenceID DESC";
        //                    isAllowed = DBUtils.ExecuteSQL(mySql);

        //                    if (Common.GetSysSettings("AutoTransferAuthorizationBSU") == "1")
        //                    {
        //                        isAllowed = false;
        //                    }
        //                    if (isAllowed == false)
        //                    {
        //                        try
        //                        {
        //                            command.Transaction = transaction;
        //                            //command.CommandTimeout = 60;
        //                            string updateCustomerTrans = "update CustomerTransactions set PaidGLReferenceID = null where PaidGLReferenceID in (select  GLReferenceID from GLReferences where  VoucherNumber >= '" + pModel.StartingRange + "' and VoucherNumber <= '" + pModel.EndingRange + "')";
        //                            command.CommandText = updateCustomerTrans;
        //                            command.ExecuteNonQuery();
        //                            command.Parameters.Clear();

        //                            command.Transaction = transaction;
        //                            string deleteCustomerTrans = "delete from CustomerTransactions where HoldGLReferenceID in (select  GLReferenceID from GLReferences where  VoucherNumber >= '" + pModel.StartingRange + "' and VoucherNumber <= '" + pModel.EndingRange + "')";
        //                            command.CommandText = deleteCustomerTrans;
        //                            command.ExecuteNonQuery();
        //                            command.Parameters.Clear();

        //                            command.Transaction = transaction;
        //                            string deleteGLTrans = "delete from GLTransactions where GLReferenceID in (select  GLReferenceID  from GLReferences where VoucherNumber >= '" + pModel.StartingRange + "' and VoucherNumber <= '" + pModel.EndingRange + "')";
        //                            command.CommandText = deleteGLTrans;
        //                            command.ExecuteNonQuery();
        //                            command.Parameters.Clear();

        //                            command.Transaction = transaction;
        //                            string deleteGLReferences = "delete from GLReferences where  GLReferenceID in (select  GLReferenceID from GLReferences where  VoucherNumber >= '" + pModel.StartingRange + "' and VoucherNumber <= '" + pModel.EndingRange + "')";
        //                            command.CommandText = deleteGLReferences;
        //                            command.ExecuteNonQuery();
        //                            command.Parameters.Clear();
        //                            int endingRange = Common.toInt(pModel.EndingRange) + 1;
        //                            int startingRange = Common.toInt(pModel.StartingRange);
        //                            int vno = endingRange - startingRange;
        //                            command.Transaction = transaction;
        //                            string updateGLReferences = "update GLReferences set  VoucherNumber = VoucherNumber - " + vno + " where VoucherNumber >='" + pModel.StartingRange + "'";
        //                            command.CommandText = updateGLReferences;
        //                            command.ExecuteNonQuery();
        //                            command.Parameters.Clear();
        //                            pModel.isPosted = true;
        //                            isDeleted = true;
        //                            transaction.Commit();
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            transaction.Rollback();
        //                            isDeleted = false;
        //                            pModel.ErrMessage = Common.GetAlertMessage(1, "Error occured.");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        transaction.Rollback();
        //                        isDeleted = false;
        //                        pModel.ErrMessage = Common.GetAlertMessage(1, "Sorry! same user can not deleted voucher posted by same user. There is one or more transactions which are posted by same user who is logged in right now.");
        //                    }
        //                }
        //                else
        //                {
        //                    isDeleted = false;
        //                    pModel.ErrMessage = Common.GetAlertMessage(1, "Please provide starting voucher no and ending voucher no.");
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                isDeleted = false;
        //                transaction.Rollback();
        //            }
        //            finally
        //            {
        //                connection.Close();
        //            }
        //        }
        //    }
        //    else
        //    {
        //        isDeleted = false;
        //        pModel.isPosted = true;
        //        pModel.ErrMessage = Common.GetAlertMessage(1, "Invalid pin code provided.");
        //    }
        //    if (isDeleted)
        //    {
        //        pModel.ErrMessage = Common.GetAlertMessage(0, "&nbsp &nbsp &nbsp Selected Vouchers deleted succcessfully.");
        //    }
        //    return View("~/Views/Postings/GeneralJournalDelete.cshtml", pModel);
        //}

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
        #endregion


        #region Update Voucher Posting Date
        public ActionResult UpdateVoucherPostingDate()
        {
            PostingViewModel postingViewModel = new PostingViewModel();
            postingViewModel.DatePosted = Common.toDateTime(SysPrefs.PostingDate);
            postingViewModel.isListPosted = true;
            postingViewModel.isPosted = true;
            postingViewModel.isBtnShow = false;
            postingViewModel.FromDate = Common.toDateTime(SysPrefs.PostingDate).AddDays(-30);
            postingViewModel.ToDate = Common.toDateTime(SysPrefs.PostingDate);
            return View("~/Views/Postings/UpdateVoucherPostingDate.cshtml", postingViewModel);
        }

        #region Single Voucher update
        [HttpPost]
        public ActionResult UpdateVoucherPostingDate(PostingViewModel pModel)
        {
            PostingViewModel postingViewModel = new PostingViewModel();
            if (!string.IsNullOrEmpty(pModel.VoucherNumber) && pModel.DatePosted.ToShortDateString() != "1/1/0001" && pModel.DatePosted != null)
            {
                string mySql = "Select GLReferenceID from GLReferences";
                ///mySql += " where VoucherNumber ='" + VoucherNumber + "'";
                mySql += " where VoucherNumber ='" + pModel.VoucherNumber + "'";
                string id = Common.toString(DBUtils.executeSqlGetSingle(mySql));
                //TransViewModel pGLTransViewModel = new TransViewModel();
                if (!string.IsNullOrEmpty(id))
                {
                    postingViewModel.cpRefID = id;
                    postingViewModel.DatePosted = pModel.DatePosted;
                    postingViewModel.HtmlTable1 = BuildDynamicTable1(Common.toInt(postingViewModel.cpRefID));
                    postingViewModel.isListPosted = false;
                    string Records = "Select VoucherNumber, AuthorizedBy, AuthorizedDate, DatePosted, PostedBy from GLReferences where GLReferenceID=" + postingViewModel.cpRefID + "";
                    DataTable transTable = DBUtils.GetDataTable(Records);
                    if (transTable != null && transTable.Rows.Count > 0)
                    {
                        foreach (DataRow RecordTable in transTable.Rows)
                        {
                            postingViewModel.AuthorizedBy = Common.GetUserFullName(Common.toString(RecordTable["AuthorizedBy"]));
                            postingViewModel.AuthorizedDate = Common.toDateTime(Common.toDateTime(RecordTable["AuthorizedDate"]).ToShortDateString());
                            postingViewModel.DatePosted = Common.toDateTime(Common.toDateTime(RecordTable["DatePosted"]).ToShortDateString());
                            postingViewModel.PostedBy = Common.GetUserFullName(Common.toString(RecordTable["PostedBy"]));
                            postingViewModel.VoucherNumber = Common.toString(RecordTable["VoucherNumber"]);
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
                postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Please provide voucher # and new posting date");
            }
            return View("~/Views/Postings/UpdateVoucherPostingDate.cshtml", postingViewModel);
        }

        public ActionResult UpdateSinglePostingDate(PostingViewModel pModel)
        {
            using (SqlConnection connection = Common.getConnection())
            {
                SqlTransaction transaction = null;
                try
                {
                    transaction = connection.BeginTransaction();
                    SqlCommand command = connection.CreateCommand();
                    ApplicationUser Profile = ApplicationUser.GetUserProfile();
                    if (Common.ValidatePinCode(Profile.Id, pModel.PinCode))
                    {
                        bool isAllowed = false;
                        //string PostedBy = DBUtils.executeSqlGetSingle("Select PostedBy from GLReferences where [glreferenceid]=" + pModel.cpRefID + "");
                        //if (Common.GetSysSettings("AutoTransferAuthorizationBSU") == "1")
                        //{
                        isAllowed = true;
                        //}
                        //else
                        //{
                        //    if (PostedBy != Profile.Id)
                        //    {
                        //        isAllowed = true;
                        //    }
                        //}
                        if (isAllowed)
                        {
                            try
                            {
                                string myRefferenceNo = Common.toString(pModel.cpRefID);
                                if (myRefferenceNo != "")
                                {
                                    command.Transaction = transaction;
                                    string SqlDeletetransaction = "update GLTransactions set addedDate=@PostingDate where GLReferenceID=@GLReferenceID";
                                    command.Parameters.AddWithValue("@GLReferenceID", pModel.cpRefID);
                                    command.Parameters.AddWithValue("@PostingDate", pModel.DatePosted);
                                    command.CommandText = SqlDeletetransaction;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();

                                    command.Transaction = transaction;
                                    string SqlDeleteVoucher = "update GLReferences set  DatePosted=@PostingDate where GLReferenceID=@GLReferenceID;";
                                    command.Parameters.AddWithValue("@GLReferenceID", myRefferenceNo);
                                    command.Parameters.AddWithValue("@PostingDate", pModel.DatePosted);
                                    command.CommandText = SqlDeleteVoucher;
                                    command.ExecuteNonQuery();
                                    command.Parameters.Clear();
                                }

                                pModel.isPosted = true;
                                pModel.ErrMessage = Common.GetAlertMessage(0, "&nbsp &nbsp &nbsp Voucher posting date updated succcessfully.");
                                transaction.Commit();
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                pModel.ErrMessage = Common.GetAlertMessage(1, "Error occured.");
                            }
                        }
                        else
                        {
                            transaction.Rollback();
                            pModel.ErrMessage = Common.GetAlertMessage(1, "Sorry! same user can not update voucher posted by same user.");
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                        pModel.isPosted = true;
                        pModel.ErrMessage = Common.GetAlertMessage(1, "Invalid pin code provided.");
                    }
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                }
                finally
                {
                    connection.Close();
                }
            }
            return View("~/Views/Postings/UpdateVoucherPostingDate.cshtml", pModel);
        }
        #endregion

        #region Range Voucher update
        [HttpPost]
        public ActionResult UpdateVoucherPostingDateRange(PostingViewModel postingViewModel)
        {
            string table = "";
            try
            {
                if (Common.toInt(postingViewModel.StartingRange) > 0 && Common.toInt(postingViewModel.EndingRange) > 0 && !string.IsNullOrEmpty(Common.toString(postingViewModel.TypeiCode)) && postingViewModel.DatePosted.ToShortDateString() != "1/1/0001" && postingViewModel.DatePosted != null)
                {
                    string mySql = "Select VoucherNumber, ReferenceNo, DatePosted, GLReferenceID,(Select Description From ConfigItemsData Where ConfigItemsData.ItemsDataCode = GLReferences.TypeiCode) As TypeiCode from GLReferences ";
                    mySql += " where GLReferences.VoucherNumber >= " + postingViewModel.StartingRange + " and GLReferences.VoucherNumber  <= " + postingViewModel.EndingRange + " and GLReferences.TypeiCode='" + postingViewModel.TypeiCode + "' order by GLReferenceID DESC";
                    DataTable dtRange = DBUtils.GetDataTable(mySql);
                    if (dtRange != null && dtRange.Rows.Count > 0)
                    {
                        table += "<table id='tblMain' border='1' width='100%' cellpadding='5'>";
                        table += "<thead>";
                        table += " <tr>";
                        //table += "<th style='background-color:#007bff; color:#FFF;'><input id='chk' type='checkbox' value=" + 1 + " onclick='checkAll(this);'></th>";
                        table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Date</th>";
                        table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Voucher #</th>";
                        table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Reference ID</th>";
                        table += "<th style='background-color:#007bff; color:#FFF; text-align:center;'>Type</th>";
                        table += "</tr>";
                        table += "<thead>";
                        table += "<tbody>";
                        foreach (DataRow dr in dtRange.Rows)
                        {
                            table += "<tr>";
                            //table += "<tr id='" + Common.toString(Item["GLReferenceID"]) + "'><td><input id='chk_" + Common.toString(Item["GLReferenceID"]) + "' type='checkbox' Class='case' value=" + Common.toString(Item["GLReferenceID"]) + " onclick='checkClick(this);'></td>";
                            table += "<td>" + dr["DatePosted"].ToString().Substring(0, 10) + "</td>";
                            table += "<td>" + dr["VoucherNumber"].ToString() + "</td>";
                            table += "<td>" + dr["GLReferenceID"].ToString() + "</td>";
                            table += "<td>" + dr["TypeiCode"].ToString() + "</td></tr>";
                        }
                        postingViewModel.isPosted = false;
                        postingViewModel.isListPosted = true;
                        postingViewModel.isBtnShow = true;
                    }
                    else
                    {
                        postingViewModel.isPosted = true;
                        postingViewModel.ErrMessage = Common.GetAlertMessage(1, "No record found.");
                    }
                }
                else
                {
                    postingViewModel.isPosted = true;
                    postingViewModel.ErrMessage = Common.GetAlertMessage(1, "Please select voucher type, starting, ending range and new posting date.");
                }
                table += "</tbody>";
                table += "</table>";
                postingViewModel.ListTable = table;
            }
            catch (Exception ex)
            {
                postingViewModel.isPosted = true;
                postingViewModel.ErrMessage = "An error occured: " + ex.ToString();
            }
            return View("~/Views/Postings/UpdateVoucherPostingDate.cshtml", postingViewModel);
        }


        [HttpPost]
        public ActionResult UpdateRangePostingDate(PostingViewModel pModel)
        {
            ApplicationUser Profile = ApplicationUser.GetUserProfile();
            bool isDeleted = false;
            if (Common.ValidatePinCode(Profile.Id, pModel.PinCode))
            {
                using (SqlConnection connection = Common.getConnection())
                {
                    SqlTransaction transaction = null;
                    try
                    {
                        transaction = connection.BeginTransaction();
                        SqlCommand command = connection.CreateCommand();
                        if (Common.toInt(pModel.StartingRange) > 0 && Common.toInt(pModel.EndingRange) > 0 && !string.IsNullOrEmpty(Common.toString(pModel.TypeiCode)) && pModel.DatePosted.ToShortDateString() != "1/1/0001" && pModel.DatePosted != null)
                        {
                            try
                            {
                                command.Transaction = transaction;
                                string SqlDeletetransaction = "update GLTransactions set addedDate=@PostingDate where GLReferenceID in (select GLReferenceID from GLReferences where  GLReferences.VoucherNumber >=@StartingRange and GLReferences.VoucherNumber <=@EndingRange and GLReferences.TypeiCode =@TypeiCode);";
                                command.Parameters.AddWithValue("@StartingRange", pModel.StartingRange);
                                command.Parameters.AddWithValue("@EndingRange", pModel.EndingRange);
                                command.Parameters.AddWithValue("@TypeiCode", pModel.TypeiCode);
                                command.Parameters.AddWithValue("@PostingDate", pModel.DatePosted);
                                command.CommandText = SqlDeletetransaction;
                                command.ExecuteNonQuery();
                                command.Parameters.Clear();

                                command.Transaction = transaction;
                                string SqlDeleteVoucher = "update GLReferences set  DatePosted=@PostingDate where GLReferences.VoucherNumber >=@StartingRange and GLReferences.VoucherNumber <=@EndingRange and GLReferences.TypeiCode =@TypeiCode;";
                                command.Parameters.AddWithValue("@StartingRange", pModel.StartingRange);
                                command.Parameters.AddWithValue("@EndingRange", pModel.EndingRange);
                                command.Parameters.AddWithValue("@TypeiCode", pModel.TypeiCode);
                                command.Parameters.AddWithValue("@PostingDate", pModel.DatePosted);
                                command.CommandText = SqlDeleteVoucher;
                                command.ExecuteNonQuery();
                                command.Parameters.Clear();
                                pModel.isPosted = true;
                                isDeleted = true;
                                transaction.Commit();
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                pModel.isPosted = true;
                                isDeleted = false;
                                pModel.ErrMessage = Common.GetAlertMessage(1, "Error occured.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        isDeleted = false;
                        //transaction.Rollback();
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }
            else
            {
                isDeleted = false;
                pModel.isPosted = true;
                pModel.ErrMessage = Common.GetAlertMessage(1, "Invalid pin code provided.");
            }
            if (isDeleted)
            {
                pModel.ErrMessage = Common.GetAlertMessage(0, "&nbsp &nbsp &nbsp Selected Vouchers updated succcessfully.");
            }
            return View("~/Views/Postings/UpdateVoucherPostingDate.cshtml", pModel);
        }

        #endregion

        #endregion


        #region Reprocess Payment
        public ActionResult ReprocessPayment()
        {
            CancelPaymentViewModel model = new CancelPaymentViewModel();
            model.hideContent = true;
            return View("~/Views/Postings/ReprocessTransactions.cshtml", model);
        }
        public JsonResult GetBuyers([DataSourceRequest] DataSourceRequest request)
        {

            ChartOfAccountsModel obj = new ChartOfAccountsModel();
            var bankAccount = _dbContext.Query<ChartOfAccountsModel>("select VendorID  as AccountCode, VendorCompany as AccountTitle from Vendors where Status=1 order by  VendorCompany").ToList();
            obj.AccountTitle = "Select new Buyer";
            obj.AccountCode = "";
            bankAccount.Insert(0, obj);
            return Json(bankAccount, JsonRequestBehavior.AllowGet);
        }
        public ActionResult ReprocessPaymentView(CancelPaymentViewModel pGLTransViewModel)
        {
            pGLTransViewModel.isPosted = false;
            pGLTransViewModel.hideContent = true;
            if (!string.IsNullOrEmpty(pGLTransViewModel.tranxRef))
            {
                string Records = "";
                if (pGLTransViewModel.Type == "1")
                {
                    Records = "Select CustomerTransactionID,PaymentNo, HoldDate, CancelledGLReferenceID, CancelledDate, SendingCountry,RecevingCountry,Address,City,TransactionID,Phone,Date,Recipient,SenderName,FC_Amount,CancellationReason,(select top 1 CustomerPrefix from Customers where Customers.CustomerID=CustomerTransactions.CustomerID ) as  CustomerPrefix,Status,PaidDate,Pounds,PaidGLReferenceID,HoldGLReferenceID  from CustomerTransactions Where PaymentNo ='" + pGLTransViewModel.tranxRef + "' and Status='Paid'";
                }
                else if (pGLTransViewModel.Type == "2")
                {
                    Records = "Select CustomerTransactionID,PaymentNo, HoldDate,CancelledGLReferenceID, CancelledDate, SendingCountry,RecevingCountry,Address,City,TransactionID,Phone,Date,Recipient,SenderName,FC_Amount,CancellationReason,(select top 1 CustomerPrefix from Customers where Customers.CustomerID=CustomerTransactions.CustomerID ) as  CustomerPrefix,Status,PaidDate,Pounds,PaidGLReferenceID,HoldGLReferenceID  from CustomerTransactions Where TransactionID='" + pGLTransViewModel.tranxRef + "' and Status='Paid'";
                }
                DataTable transTable = DBUtils.GetDataTable(Records);
                if (transTable != null)
                {
                    if (transTable.Rows.Count > 0)
                    {
                        DataRow drTransaction = transTable.Rows[0];
                        if (drTransaction != null)
                        {
                            pGLTransViewModel.CustomerTransactionID = Common.toString(drTransaction["CustomerTransactionID"]);
                            pGLTransViewModel.TransactionID = Common.toString(drTransaction["TransactionID"]);
                            if (Common.toString(drTransaction["Status"]).Trim() != "Cancelled")
                            {
                                if (Common.toString(drTransaction["CancelledGLReferenceID"]) == "" || Common.toString(drTransaction["CancelledGLReferenceID"]) == "0")
                                {
                                    string getVoucherNo = "SELECT distinct GLReferences.VoucherNumber, GLTransactions.GLReferenceID FROM GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID where TransactionNumber = '" + Common.toString(drTransaction["PaymentNo"]) + "'";
                                    DataTable dtVouchers = DBUtils.GetDataTable(getVoucherNo);
                                    if (dtVouchers != null && dtVouchers.Rows.Count > 0)
                                    {
                                        foreach (DataRow drVoucher in dtVouchers.Rows)
                                        {
                                            pGLTransViewModel.VoucherNo += Common.toString(drVoucher["VoucherNumber"]) + ", ";
                                            pGLTransViewModel.cpGLReferenceID = Common.toString(drVoucher["GLReferenceID"]);
                                        }
                                    }
                                    string getAccountName = "select AccountName, AccountCode, CurrencyISOCode from GLChartOfAccounts where Prefix = '" + Common.toString(drTransaction["CustomerPrefix"]) + "'";
                                    DataTable getHoldAcc = DBUtils.GetDataTable(getAccountName);
                                    if (getHoldAcc != null && getHoldAcc.Rows.Count > 0)
                                    {
                                        DataRow drRow = getHoldAcc.Rows[0];
                                        pGLTransViewModel.CurrencyISOCode = Common.toString(drRow["CurrencyISOCode"]);
                                        pGLTransViewModel.AccountName = Common.toString(drRow["AccountName"]);
                                        pGLTransViewModel.HoldAccount = Common.toString(drRow["AccountCode"]);

                                        pGLTransViewModel.CurrencyType = DBUtils.executeSqlGetSingle("Select CurrenyType from ListCurrencies Where CurrencyISOCode = '" + pGLTransViewModel.CurrencyISOCode + "'");
                                        pGLTransViewModel.AgentCurrRate = DBUtils.executeSqlGetSingle("Select ExchangeRate from ExchangeRates Where CurrencyISOCode='" + pGLTransViewModel.CurrencyISOCode + "' ");
                                    }
                                    pGLTransViewModel.CustomerPrefix = Common.toString(drTransaction["CustomerPrefix"]);
                                    pGLTransViewModel.Phone = Common.toString(drTransaction["Phone"]);
                                    pGLTransViewModel.PaymentNo = Common.toString(drTransaction["PaymentNo"]);
                                    pGLTransViewModel.InvoiceNo = Common.toString(drTransaction["TransactionID"]);
                                    pGLTransViewModel.Date = Common.toDateTime(drTransaction["Date"]);
                                    pGLTransViewModel.CancelledDate = Common.toDateTime(drTransaction["CancelledDate"]);
                                    pGLTransViewModel.Recipient = Common.toString(drTransaction["Recipient"]) + ", " + Common.toString(drTransaction["Address"] + "-" + Common.toString(drTransaction["City"]));
                                    pGLTransViewModel.SenderName = Common.toString(drTransaction["SenderName"] + "," + drTransaction["SendingCountry"]);
                                    pGLTransViewModel.FC_Amount = Common.toDecimal(drTransaction["FC_Amount"]);
                                    pGLTransViewModel.CancellationReason = Common.toString(drTransaction["CancellationReason"]);
                                    pGLTransViewModel.Status = Common.toString(drTransaction["Status"]);
                                    pGLTransViewModel.PaidDate = Common.toDateTime(drTransaction["PaidDate"]);
                                    pGLTransViewModel.HoldDate = Common.toDateTime(drTransaction["HoldDate"]);
                                    pGLTransViewModel.Pounds = Common.toDecimal(drTransaction["Pounds"]);
                                    pGLTransViewModel.PaidGLReferenceID = Common.toString(drTransaction["PaidGLReferenceID"]);
                                    pGLTransViewModel.HoldGLReferenceID = Common.toString(drTransaction["HoldGLReferenceID"]);
                                    pGLTransViewModel = ReprocessTranxBuildTable(pGLTransViewModel);
                                    pGLTransViewModel.isPosted = true;
                                }
                                else
                                {
                                    pGLTransViewModel.CancellationMessage = "&nbsp &nbsp &nbsp &nbsp <h4>Transaction already cancelled and pending for authorization.</h4>";
                                }
                            }
                            else
                            {
                                pGLTransViewModel.CancellationMessage = "&nbsp &nbsp &nbsp &nbsp <h4>Transaction already cancelled.</h4>";
                            }

                        }
                        else
                        {
                            pGLTransViewModel.ErrMessage = Common.GetAlertMessage(1, "Transaction not found.");
                        }
                    }
                    else
                    {
                        pGLTransViewModel.ErrMessage = Common.GetAlertMessage(1, "No data found against  <b>" + pGLTransViewModel.tranxRef + " </b> Invoice/Payment No.");
                    }

                }
                else
                {
                    pGLTransViewModel.ErrMessage = Common.GetAlertMessage(1, "No data found against  <b>" + pGLTransViewModel.tranxRef + "</b>  Invoice/Payment No.");
                }
            }
            else
            {
                pGLTransViewModel.ErrMessage = Common.GetAlertMessage(1, "Invalid Payment#/Invoice#");
            }
            return View("~/Views/Postings/ReprocessTransactions.cshtml", pGLTransViewModel);
        }
        protected CancelPaymentViewModel ReprocessTranxBuildTable(CancelPaymentViewModel model)
        {
            string table = "";
            decimal cpTotalDebit = 0m;
            decimal cpTotalCredit = 0m;
            string FC = "0.00";
            string LC = "0.00";
            int iRowNumber = 0;
            decimal LcAmount = 0m;
            decimal FcAmount = 0m;
            string tranxQury = "";
            if (Common.toString(model.PaidGLReferenceID) != "")
            {
                tranxQury = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.PaidGLReferenceID + "' order by LocalCurrencyAmount Desc";
                DataTable objGLTransView = DBUtils.GetDataTable(tranxQury);
                if (objGLTransView != null && objGLTransView.Rows.Count > 0)
                {

                    foreach (DataRow MyRow in objGLTransView.Rows)
                    {
                        string sqlListCurrencies = "select HoldAccountCode from ListCurrencies where CurrencyISOCode='" + Common.toString(MyRow["ForeignCurrencyISOCode"]).Trim() + "'";
                        string myHoldAccountCode = DBUtils.executeSqlGetSingle(sqlListCurrencies);

                        if (model.Status == "Paid" && myHoldAccountCode == Common.toString(MyRow["AccountCode"]).Trim())
                        {
                            string myAgentAccountCode = "";
                            string myAgentAccountName = "";
                            string sqlGLChartOfAccounts = "select AccountCode,AccountName from  GLChartOfAccounts where prefix='" + model.CustomerPrefix + "'";
                            DataTable dtAgentNameCode = DBUtils.GetDataTable(sqlGLChartOfAccounts);
                            if (dtAgentNameCode != null && dtAgentNameCode.Rows.Count > 0)
                            {
                                myAgentAccountCode = dtAgentNameCode.Rows[0]["AccountCode"].ToString();
                                myAgentAccountName = dtAgentNameCode.Rows[0]["AccountName"].ToString();
                            }
                            string sqlAgentTransaction = "select AccountCode,ForeignCurrencyISOCode, GLReferenceID,LocalCurrencyAmount,ForeignCurrencyAmount,ExchangeRate,BaseAmount,Memo,(Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName from GLTransactions where GLReferenceID='" + model.PaidGLReferenceID + "' and AccountCode ='" + myHoldAccountCode + "' and LocalCurrencyAmount='" + MyRow["LocalCurrencyAmount"].ToString() + "'";
                            DataTable dtAgentTransaction = DBUtils.GetDataTable(sqlAgentTransaction);
                            if (dtAgentTransaction != null && dtAgentTransaction.Rows.Count > 0)
                            {
                                DataRow drAgentTransaction = dtAgentTransaction.Rows[0];
                                if (!string.IsNullOrEmpty(drAgentTransaction["ForeignCurrencyAmount"].ToString()))
                                {
                                    FC = drAgentTransaction["ForeignCurrencyAmount"].ToString();
                                }
                                table += "<tr><td style='padding: 2px' valign='top'><strong>" + drAgentTransaction["AccountCode"].ToString() + "-" + drAgentTransaction["AccountName"].ToString() + "</strong>" + "<br/>" + drAgentTransaction["Memo"].ToString() + "<br/>" + drAgentTransaction["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAgentTransaction["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAgentTransaction["GLReferenceID"].ToString() + "</td>";
                                table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAgentTransaction["ForeignCurrencyISOCode"].ToString() + "</td>";
                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";

                                if (!string.IsNullOrEmpty(drAgentTransaction["LocalCurrencyAmount"].ToString()))
                                {
                                    LC = drAgentTransaction["LocalCurrencyAmount"].ToString();
                                }
                                LcAmount = -1 * Convert.ToDecimal(LC);

                                if (LcAmount < 0)
                                {
                                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "</td>";
                                    cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                                    model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                                }
                                else
                                {
                                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                    cpTotalDebit = cpTotalDebit + LcAmount;
                                    model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                                }
                                table += "</tr>";

                                #region In case of Paid 
                                if (model.Status == "Paid")
                                {
                                    #region Admin charges
                                    //if (Common.toString(SysPrefs.TransactionAdminCommissionAccount) != "" && Common.toString(model.HoldGLReferenceID) != "")
                                    //{
                                    //    string sqlAdminCharges = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode='" + SysPrefs.TransactionAdminCommissionAccount + "' order by LocalCurrencyAmount Desc";
                                    //    DataTable ObjAdminCharges = DBUtils.GetDataTable(sqlAdminCharges);
                                    //    if (ObjAdminCharges != null && ObjAdminCharges.Rows.Count > 0)
                                    //    {
                                    //        iRowNumber = 0;
                                    //        foreach (DataRow drAdminCharges in ObjAdminCharges.Rows)
                                    //        {
                                    //            //Debit Agent account charges
                                    //            FC = "0.00";
                                    //            if (!string.IsNullOrEmpty(drAdminCharges["ForeignCurrencyAmount"].ToString()))
                                    //            {
                                    //                FC = drAdminCharges["ForeignCurrencyAmount"].ToString();
                                    //            }
                                    //            FcAmount = Math.Abs(Convert.ToDecimal(FC));
                                    //            table += "<tr><td style='padding: 2px' valign='top'><strong>" + drAdminCharges["AccountCode"].ToString() + "-" + drAdminCharges["AccountName"].ToString() + "</strong>" + "<br/>" + drAdminCharges["Memo"].ToString() + "<br/>" + drAdminCharges["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAdminCharges["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAdminCharges["GLReferenceID"].ToString() + "</td>";
                                    //            table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAdminCharges["ForeignCurrencyISOCode"].ToString() + "</td>";
                                    //            table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";
                                    //            LC = "0.00";
                                    //            if (!string.IsNullOrEmpty(drAdminCharges["LocalCurrencyAmount"].ToString()))
                                    //            {
                                    //                LC = drAdminCharges["LocalCurrencyAmount"].ToString();
                                    //            }
                                    //            LcAmount = -1 * Convert.ToDecimal(LC);

                                    //            if (LcAmount < 0)
                                    //            {
                                    //                table += "<td style='padding: 2px'>&nbsp;</td>";
                                    //                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat((-1 * LcAmount)) + "</td>";
                                    //                cpTotalCredit = cpTotalCredit + (-1 * LcAmount);
                                    //                model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                                    //            }
                                    //            else
                                    //            {
                                    //                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                    //                table += "<td style='padding: 2px'>&nbsp;</td>";
                                    //                cpTotalDebit = cpTotalDebit + LcAmount;
                                    //                model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                                    //            }
                                    //            table += "</tr>";
                                    //            iRowNumber++;

                                    //            //Credit Agent account charges

                                    //            FC = "0.00";
                                    //            if (!string.IsNullOrEmpty(drAdminCharges["ForeignCurrencyAmount"].ToString()))
                                    //            {
                                    //                FC = drAdminCharges["ForeignCurrencyAmount"].ToString();
                                    //            }
                                    //            FcAmount = Math.Abs(Convert.ToDecimal(FC));
                                    //            table += "<tr><td style='padding: 2px' valign='top'><strong>" + myAgentAccountCode + "-" + myAgentAccountName + "</strong>" + "<br/>" + drAdminCharges["Memo"].ToString() + "<br/>" + drAdminCharges["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAdminCharges["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAdminCharges["GLReferenceID"].ToString() + "</td>";
                                    //            table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAdminCharges["ForeignCurrencyISOCode"].ToString() + "</td>";
                                    //            table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";
                                    //            LC = "0.00";
                                    //            if (!string.IsNullOrEmpty(drAdminCharges["LocalCurrencyAmount"].ToString()))
                                    //            {
                                    //                LC = drAdminCharges["LocalCurrencyAmount"].ToString();
                                    //            }
                                    //            LcAmount = 1 * Convert.ToDecimal(LC);

                                    //            if (LcAmount < 0)
                                    //            {
                                    //                table += "<td style='padding: 2px'>&nbsp;</td>";
                                    //                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat((-1 * LcAmount)) + "</td>";
                                    //                cpTotalCredit = cpTotalCredit + (-1 * LcAmount);
                                    //                model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                                    //            }
                                    //            else
                                    //            {
                                    //                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                    //                table += "<td style='padding: 2px'>&nbsp;</td>";
                                    //                cpTotalDebit = cpTotalDebit + LcAmount;
                                    //                model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                                    //            }
                                    //            table += "</tr>";
                                    //            iRowNumber++;
                                    //        }
                                    //    }
                                    //}
                                    #endregion

                                    #region Agent charges
                                    //if (Common.toString(SysPrefs.TransactionAgentCommissionAccount) != "" && Common.toString(model.HoldGLReferenceID) != "")
                                    //{
                                    //    if (Common.toString(SysPrefs.TransactionAdminCommissionAccount) != Common.toString(SysPrefs.TransactionAgentCommissionAccount))
                                    //    {
                                    //        string sqlAgentCharges = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode='" + SysPrefs.TransactionAgentCommissionAccount + "' order by LocalCurrencyAmount Desc";
                                    //        DataTable objAgentCharges = DBUtils.GetDataTable(sqlAgentCharges);
                                    //        if (objAgentCharges != null && objAgentCharges.Rows.Count > 0)
                                    //        {
                                    //            foreach (DataRow drAgentCharges in objAgentCharges.Rows)
                                    //            {
                                    //                //Debit Agent account charges
                                    //                if (!string.IsNullOrEmpty(drAgentCharges["ForeignCurrencyAmount"].ToString()))
                                    //                {
                                    //                    FC = drAgentCharges["ForeignCurrencyAmount"].ToString();
                                    //                }
                                    //                FcAmount = Math.Abs(Convert.ToDecimal(FC));
                                    //                table += "<tr><td style='padding: 2px' valign='top'><strong>" + drAgentCharges["AccountCode"].ToString() + "-" + drAgentCharges["AccountName"].ToString() + "</strong>" + "<br/>" + drAgentCharges["Memo"].ToString() + "<br/>" + drAgentCharges["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAgentCharges["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAgentCharges["GLReferenceID"].ToString() + "</td>";
                                    //                table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAgentCharges["ForeignCurrencyISOCode"].ToString() + "</td>";
                                    //                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";

                                    //                if (!string.IsNullOrEmpty(drAgentCharges["LocalCurrencyAmount"].ToString()))
                                    //                {
                                    //                    LC = drAgentCharges["LocalCurrencyAmount"].ToString();
                                    //                }
                                    //                LcAmount = -1 * Convert.ToDecimal(LC);

                                    //                if (LcAmount < 0)
                                    //                {
                                    //                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                    //                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat((-1 * LcAmount)) + "</td>";
                                    //                    cpTotalCredit = cpTotalCredit + (-1 * LcAmount);
                                    //                    model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                                    //                }
                                    //                else
                                    //                {
                                    //                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                    //                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                    //                    cpTotalDebit = cpTotalDebit + LcAmount;
                                    //                    model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                                    //                }
                                    //                table += "</tr>";
                                    //                iRowNumber++;

                                    //                //Credit Agent account charges

                                    //                if (!string.IsNullOrEmpty(drAgentCharges["ForeignCurrencyAmount"].ToString()))
                                    //                {
                                    //                    FC = drAgentCharges["ForeignCurrencyAmount"].ToString();
                                    //                }
                                    //                FcAmount = Math.Abs(Convert.ToDecimal(FC));
                                    //                table += "<tr><td style='padding: 2px' valign='top'><strong>" + myAgentAccountCode + "-" + myAgentAccountName + "</strong>" + "<br/>" + drAgentCharges["Memo"].ToString() + "<br/>" + drAgentCharges["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + drAgentCharges["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + drAgentCharges["GLReferenceID"].ToString() + "</td>";
                                    //                table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + drAgentCharges["ForeignCurrencyISOCode"].ToString() + "</td>";
                                    //                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";

                                    //                if (!string.IsNullOrEmpty(drAgentCharges["LocalCurrencyAmount"].ToString()))
                                    //                {
                                    //                    LC = drAgentCharges["LocalCurrencyAmount"].ToString();
                                    //                }
                                    //                LcAmount = 1 * Convert.ToDecimal(LC);

                                    //                if (LcAmount < 0)
                                    //                {
                                    //                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                    //                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat((-1 * LcAmount)) + "</td>";
                                    //                    cpTotalCredit = cpTotalCredit + (-1 * LcAmount);
                                    //                    model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                                    //                }
                                    //                else
                                    //                {
                                    //                    table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                    //                    table += "<td style='padding: 2px'>&nbsp;</td>";
                                    //                    cpTotalDebit = cpTotalDebit + LcAmount;
                                    //                    model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                                    //                }
                                    //                table += "</tr>";
                                    //                iRowNumber++;
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    #endregion
                                }
                                #endregion

                            }
                            iRowNumber++;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(MyRow["ForeignCurrencyAmount"].ToString()))
                            {
                                FC = MyRow["ForeignCurrencyAmount"].ToString();
                            }
                            FcAmount = Math.Abs(Convert.ToDecimal(FC));
                            table += "<tr><td style='padding: 2px' valign='top'><strong>" + MyRow["AccountCode"].ToString() + "-" + MyRow["AccountName"].ToString() + "</strong>" + "<br/>" + MyRow["Memo"].ToString() + "<br/>" + MyRow["BaseAmount"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Exchange Rate:" + MyRow["ExchangeRate"].ToString() + " &nbsp; &nbsp; &nbsp;" + "Ref: " + MyRow["GLReferenceID"].ToString() + "</td>";
                            table += "<td valign='top' align='center' style='padding: 2px;border-left:transparent 1px solid;'>" + MyRow["ForeignCurrencyISOCode"].ToString() + "</td>";
                            table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(FcAmount) + "</td>";

                            if (!string.IsNullOrEmpty(MyRow["LocalCurrencyAmount"].ToString()))
                            {
                                LC = MyRow["LocalCurrencyAmount"].ToString();
                            }
                            LcAmount = -1 * Convert.ToDecimal(LC);

                            if (LcAmount < 0)
                            {
                                table += "<td style='padding: 2px'>&nbsp;</td>";
                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(-1 * LcAmount) + "</td>";
                                cpTotalCredit = cpTotalCredit + -1 * LcAmount;
                                model.TotalCredit = DisplayUtils.GetSystemAmountFormat(cpTotalCredit);
                            }
                            else
                            {
                                table += "<td style='padding: 2px' valign='top' align='right'>" + DisplayUtils.GetSystemAmountFormat(LcAmount) + "</td>";
                                table += "<td style='padding: 2px'>&nbsp;</td>";
                                cpTotalDebit = cpTotalDebit + LcAmount;
                                model.TotalDebit = DisplayUtils.GetSystemAmountFormat(cpTotalDebit);
                            }
                            table += "</tr>";
                            iRowNumber++;
                        }

                    }

                }
                else
                {
                    model.Table = Common.GetAlertMessage(1, "Data not found.");
                }
            }
            else
            {
                model.Table = Common.GetAlertMessage(1, "Data not found.");
            }
            table += "</tbody>";
            model.Table = table;
            return model;
        }
        public ActionResult ReprocessSubmit(CancelPaymentViewModel model)
        {
            string StrQury = "";
            model.hideContent = true;
            bool isValid = true;
            if (Common.toString(model.GeneralAccount).Trim() == "" && Common.toString(model.BuyerRate).Trim() == "" && Common.toString(model.BuyerRateSC).Trim() == "")
            {
                isValid = false;
                model.ErrMessage = Common.GetAlertMessage(1, "Incomplete data. please provide new Buyer, BuyerRate and BuyerRateSC.");
            }
            if (Common.toString(model.CustomerTransactionID).Trim() == "" && Common.toString(model.TransactionID).Trim() == "")
            {
                isValid = false;
                model.ErrMessage = Common.GetAlertMessage(1, "Incomplete data. TransactionID not found.");
            }
            string sqlTranx = "select VendorCurrency from Vendors where VendorID=" + model.GeneralAccount + "";
            string PayoutCurrency = Common.toString(DBUtils.executeSqlGetSingle(sqlTranx));
            if (PayoutCurrency != "")
            {
                if (isValid)
                {
                    decimal BuyerRate = Common.toDecimal(model.BuyerRate);
                    if (Common.toString(PayoutCurrency).Trim() != "" && BuyerRate > 0)
                    {
                        decimal MinRateLimit = 0m;
                        decimal MiaxRateLimit = 0m;
                        DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + PayoutCurrency + "'");
                        if (objCurrency != null)
                        {
                            if (objCurrency.Rows.Count > 0)
                            {
                                DataRow drCurrency = objCurrency.Rows[0];
                                MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                            }
                            if (BuyerRate < MinRateLimit || BuyerRate > MiaxRateLimit)
                            {
                                isValid = false;
                                model.ErrMessage = Common.GetAlertMessage(1, "Rate can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>");
                            }
                        }
                    }
                    else
                    {
                        isValid = false;
                        model.ErrMessage = Common.GetAlertMessage(1, "Buyer rate should be > 0<br>");
                    }

                }

                if (isValid)
                {
                    decimal BuyerRateSC = Common.toDecimal(model.BuyerRateSC);
                    if (Common.toString(PayoutCurrency).Trim() != "" && BuyerRateSC > 0)
                    {
                        decimal MinRateLimit = 0m;
                        decimal MiaxRateLimit = 0m;
                        DataTable objCurrency = DBUtils.GetDataTable("Select MinRateLimit, MaxRateLimit from ListCurrencies where CurrencyISOCode='" + PayoutCurrency + "'");
                        if (objCurrency != null)
                        {
                            if (objCurrency.Rows.Count > 0)
                            {
                                DataRow drCurrency = objCurrency.Rows[0];
                                MinRateLimit = Common.toDecimal(drCurrency["MinRateLimit"]);
                                MiaxRateLimit = Common.toDecimal(drCurrency["MaxRateLimit"]);
                            }
                            if (BuyerRateSC < MinRateLimit || BuyerRateSC > MiaxRateLimit)
                            {
                                isValid = false;
                                model.ErrMessage = Common.GetAlertMessage(1, "BuyerRateSC can not be greater then max rate [" + MiaxRateLimit.ToString() + "] or less then min rate [" + MinRateLimit.ToString() + "]<br>");
                            }
                        }
                    }
                    else
                    {
                        isValid = false;
                        model.ErrMessage = Common.GetAlertMessage(1, "BuyerRateSC rate should be > 0<br>");
                    }
                }
            }
            else
            {
                isValid = false;
                model.ErrMessage = Common.GetAlertMessage(1, "Incomplete data. PayoutCurrency not found.");
            }
            if (isValid)
            {
                decimal FcAmount = 0.0m;
                decimal LcAmount = 0.0m;
                model.isPosted = false;
                if (!string.IsNullOrEmpty(model.cpGLReferenceID))
                {
                    if (Common.toString(model.Status).Trim() != "Cancelled" && Common.toString(model.CustomerPrefix).Trim() != "")
                    {
                        DataTable dtAdminCharges = new DataTable();
                        DataTable dtAgentCharges = new DataTable();

                        if (model.Status == "Paid")
                        {
                            StrQury = "select TransactionID, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, GLReferenceID  from GLTransactions where GLReferenceID='" + model.PaidGLReferenceID + "' and AccountTypeiCode='22PAIDENTRY' order by LocalCurrencyAmount Desc";

                            DataTable objGLTransView = DBUtils.GetDataTable(StrQury);
                            if (objGLTransView != null && objGLTransView.Rows.Count > 0)
                            {
                                Guid guid = Guid.NewGuid();
                                string cpUserRefNo = guid.ToString();

                                #region Reversal of existing voucher

                                //// string myConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
                                using (SqlConnection connection = Common.getConnection())
                                {
                                    SqlCommand command = connection.CreateCommand();
                                    SqlTransaction transaction = null;
                                    try
                                    {
                                        // BeginTransaction() Requires Open Connection
                                        //// connection.Open();
                                        transaction = connection.BeginTransaction();
                                        var myPostingDate = SysPrefs.PostingDate;
                                        string myDefaultCurrency = SysPrefs.DefaultCurrency;
                                        foreach (DataRow MyRow in objGLTransView.Rows)
                                        {
                                            command.Parameters.Clear();
                                            string myHoldAccountCode = "";
                                            string sqlListCurrencies = "select HoldAccountCode from ListCurrencies where CurrencyISOCode=@CurrencyISOCode";
                                            command.Transaction = transaction;
                                            command.Parameters.AddWithValue("@CurrencyISOCode", Common.toString(MyRow["ForeignCurrencyISOCode"]).Trim());
                                            command.CommandText = sqlListCurrencies;
                                            if (command.ExecuteScalar() != null)
                                            {
                                                myHoldAccountCode = command.ExecuteScalar().ToString();
                                            }
                                            command.Parameters.Clear();
                                            if (myHoldAccountCode == Common.toString(MyRow["AccountCode"]).Trim())
                                            {
                                                //model.CustomerPrefix
                                                string myAgentCurrencyIsoCode = "";
                                                string myAgentAccountCode = "";
                                                string sqlGLChartOfAccounts = "select CurrencyISOCode, AccountCode from  GLChartOfAccounts where prefix=@prefix";
                                                command.Transaction = transaction;
                                                command.Parameters.AddWithValue("@prefix", model.CustomerPrefix);
                                                command.CommandText = sqlGLChartOfAccounts;

                                                DataTable dtCoA = new DataTable();
                                                SqlDataAdapter DataAdaptter = new SqlDataAdapter(command);
                                                DataAdaptter.Fill(dtCoA);
                                                if (dtCoA != null)
                                                {
                                                    if (dtCoA.Rows.Count > 0)
                                                    {
                                                        myAgentCurrencyIsoCode = Common.toString(dtCoA.Rows[0]["CurrencyISOCode"]);
                                                        myAgentAccountCode = Common.toString(dtCoA.Rows[0]["AccountCode"]);
                                                    }
                                                }
                                                command.Parameters.Clear();
                                                //string sqlQury = "select * from GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode ='" + myAgentAccountCode + "' and LocalCurrencyAmount='" + MyRow["LocalCurrencyAmount"].ToString() + "'";
                                                string sqlQury = "select * from GLTransactions where GLReferenceID='" + model.PaidGLReferenceID + "' and AccountCode ='" + myHoldAccountCode + "' and LocalCurrencyAmount='" + MyRow["LocalCurrencyAmount"].ToString() + "'";
                                                DataTable dtAgent = DBUtils.GetDataTable(sqlQury);
                                                if (dtAgent != null && dtAgent.Rows.Count > 0)
                                                {
                                                    ///TransactionNumber = dtAgent.Rows[0]["TransactionNumber"].ToString();
                                                    string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                    command.Transaction = transaction;
                                                    command.CommandTimeout = 60;
                                                    string accCode = dtAgent.Rows[0]["AccountCode"].ToString();
                                                    FcAmount = -1 * Convert.ToDecimal(dtAgent.Rows[0]["ForeignCurrencyAmount"].ToString());
                                                    LcAmount = -1 * Convert.ToDecimal(dtAgent.Rows[0]["LocalCurrencyAmount"].ToString());
                                                    command.Parameters.AddWithValue("@AccountCode", dtAgent.Rows[0]["AccountCode"].ToString());
                                                    command.Parameters.AddWithValue("@GLTransactionTypeiCode", dtAgent.Rows[0]["AccountTypeiCode"].ToString());
                                                    command.Parameters.AddWithValue("@Memo", Common.toString(dtAgent.Rows[0]["Memo"]) + " (cancellation)");
                                                    command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                    command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                    command.Parameters.AddWithValue("@ExchangeRate", dtAgent.Rows[0]["ExchangeRate"].ToString());
                                                    command.Parameters.AddWithValue("@ForeignCurrencyISOCode", dtAgent.Rows[0]["ForeignCurrencyISOCode"].ToString());
                                                    command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                    command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                    command.CommandText = sqlMyEntry;
                                                    command.ExecuteNonQuery();
                                                }

                                                #region Revert Admin & agent charges already paid
                                                //if (Common.toString(SysPrefs.TransactionAdminCommissionAccount) != "" && Common.toString(model.HoldGLReferenceID) != "")
                                                //{
                                                //    string sqlAdminCharges = "select TransactionID,TransactionNumber, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode='" + SysPrefs.TransactionAdminCommissionAccount + "' order by LocalCurrencyAmount Desc";
                                                //    dtAdminCharges = DBUtils.GetDataTable(sqlAdminCharges);
                                                //    if (dtAdminCharges != null)
                                                //    {
                                                //        if (dtAdminCharges.Rows.Count > 0)
                                                //        {
                                                //            foreach (DataRow drAdminCharges in dtAdminCharges.Rows)
                                                //            {
                                                //                command.Parameters.Clear();
                                                //                string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                //                FcAmount = -1 * Convert.ToDecimal(drAdminCharges["ForeignCurrencyAmount"].ToString());
                                                //                LcAmount = -1 * Convert.ToDecimal(drAdminCharges["LocalCurrencyAmount"].ToString());
                                                //                command.Parameters.AddWithValue("@AccountCode", drAdminCharges["AccountCode"].ToString());
                                                //                command.Parameters.AddWithValue("@GLTransactionTypeiCode", drAdminCharges["AccountTypeiCode"].ToString());
                                                //                command.Parameters.AddWithValue("@Memo", Common.toString(drAdminCharges["Memo"]) + " (cancellation)");
                                                //                command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                //                command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                //                command.Parameters.AddWithValue("@ExchangeRate", drAdminCharges["ExchangeRate"].ToString());
                                                //                command.Parameters.AddWithValue("@ForeignCurrencyISOCode", drAdminCharges["ForeignCurrencyISOCode"].ToString());
                                                //                command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                //                command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                //                command.CommandText = sqlMyEntry;
                                                //                command.ExecuteNonQuery();
                                                //                command.Parameters.Clear();

                                                //                sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                //                FcAmount = 1 * Convert.ToDecimal(drAdminCharges["ForeignCurrencyAmount"].ToString());
                                                //                LcAmount = 1 * Convert.ToDecimal(drAdminCharges["LocalCurrencyAmount"].ToString());
                                                //                command.Parameters.AddWithValue("@AccountCode", myAgentAccountCode);
                                                //                command.Parameters.AddWithValue("@GLTransactionTypeiCode", drAdminCharges["AccountTypeiCode"].ToString());
                                                //                command.Parameters.AddWithValue("@Memo", Common.toString(drAdminCharges["Memo"]) + " (cancellation)");
                                                //                command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                //                command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                //                command.Parameters.AddWithValue("@ExchangeRate", drAdminCharges["ExchangeRate"].ToString());
                                                //                command.Parameters.AddWithValue("@ForeignCurrencyISOCode", drAdminCharges["ForeignCurrencyISOCode"].ToString());
                                                //                command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                //                command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                //                command.CommandText = sqlMyEntry;
                                                //                command.ExecuteNonQuery();
                                                //                command.Parameters.Clear();
                                                //            }
                                                //        }
                                                //    }
                                                //}
                                                //if (Common.toString(SysPrefs.TransactionAgentCommissionAccount) != "" && Common.toString(model.HoldGLReferenceID) != "")
                                                //{
                                                //    if (Common.toString(SysPrefs.TransactionAdminCommissionAccount) != Common.toString(SysPrefs.TransactionAgentCommissionAccount))
                                                //    {
                                                //        string sqlAgentCharges = "select TransactionID,TransactionNumber, ExchangeRate, ForeignCurrencyISOCode, GLReferenceID, Type, AccountTypeiCode, AccountCode, (Select AccountName from GLChartOfAccounts Where GLChartOfAccounts.AccountCode = GLTransactions.AccountCode)  as AccountName, Memo, BaseAmount, LocalCurrencyAmount, ForeignCurrencyAmount, GLPersonTypeiCode, PersonID, DimensionId, Dimension2Id, AddedBy, AddedDate, (Select GLReferenceID from GLReferences Where GLReferences.GLReferenceID = GLTransactions.GLReferenceID) as GLReferenceID From GLTransactions where GLReferenceID='" + model.HoldGLReferenceID + "' and AccountCode='" + SysPrefs.TransactionAgentCommissionAccount + "' order by LocalCurrencyAmount Desc";
                                                //        dtAgentCharges = DBUtils.GetDataTable(sqlAgentCharges);
                                                //        if (dtAgentCharges != null)
                                                //        {

                                                //            if (dtAgentCharges.Rows.Count > 0)
                                                //            {
                                                //                foreach (DataRow drAgentCharges in dtAgentCharges.Rows)
                                                //                {
                                                //                    command.Parameters.Clear();
                                                //                    string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                //                    FcAmount = -1 * Convert.ToDecimal(drAgentCharges["ForeignCurrencyAmount"].ToString());
                                                //                    LcAmount = -1 * Convert.ToDecimal(drAgentCharges["LocalCurrencyAmount"].ToString());
                                                //                    command.Parameters.AddWithValue("@AccountCode", drAgentCharges["AccountCode"].ToString());
                                                //                    command.Parameters.AddWithValue("@GLTransactionTypeiCode", drAgentCharges["AccountTypeiCode"].ToString());
                                                //                    command.Parameters.AddWithValue("@Memo", Common.toString(drAgentCharges["Memo"]) + " (cancellation)");
                                                //                    command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                //                    command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                //                    command.Parameters.AddWithValue("@ExchangeRate", drAgentCharges["ExchangeRate"].ToString());
                                                //                    command.Parameters.AddWithValue("@ForeignCurrencyISOCode", drAgentCharges["ForeignCurrencyISOCode"].ToString());
                                                //                    command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                //                    command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                //                    command.CommandText = sqlMyEntry;
                                                //                    command.ExecuteNonQuery();
                                                //                    command.Parameters.Clear();

                                                //                    sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                //                    FcAmount = 1 * Convert.ToDecimal(drAgentCharges["ForeignCurrencyAmount"].ToString());
                                                //                    LcAmount = 1 * Convert.ToDecimal(drAgentCharges["LocalCurrencyAmount"].ToString());
                                                //                    command.Parameters.AddWithValue("@AccountCode", myAgentAccountCode);
                                                //                    command.Parameters.AddWithValue("@GLTransactionTypeiCode", drAgentCharges["AccountTypeiCode"].ToString());
                                                //                    command.Parameters.AddWithValue("@Memo", Common.toString(drAgentCharges["Memo"]) + " (cancellation)");
                                                //                    command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                //                    command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                //                    command.Parameters.AddWithValue("@ExchangeRate", drAgentCharges["ExchangeRate"].ToString());
                                                //                    command.Parameters.AddWithValue("@ForeignCurrencyISOCode", drAgentCharges["ForeignCurrencyISOCode"].ToString());
                                                //                    command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                //                    command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                //                    command.CommandText = sqlMyEntry;
                                                //                    command.ExecuteNonQuery();
                                                //                    command.Parameters.Clear();
                                                //                }
                                                //            }
                                                //        }
                                                //    }
                                                //}
                                                #endregion

                                            }
                                            else
                                            {

                                                string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                                command.Transaction = transaction;
                                                command.CommandTimeout = 60;
                                                string accCode = MyRow["AccountCode"].ToString();
                                                FcAmount = -1 * Convert.ToDecimal(MyRow["ForeignCurrencyAmount"].ToString());
                                                LcAmount = -1 * Convert.ToDecimal(MyRow["LocalCurrencyAmount"].ToString());
                                                command.Parameters.AddWithValue("@AccountCode", MyRow["AccountCode"].ToString());
                                                command.Parameters.AddWithValue("@GLTransactionTypeiCode", MyRow["AccountTypeiCode"].ToString());
                                                command.Parameters.AddWithValue("@Memo", Common.toString(MyRow["Memo"]) + " (cancellation)");
                                                command.Parameters.AddWithValue("@ForeignCurrencyAmount", FcAmount);
                                                command.Parameters.AddWithValue("@LocalCurrencyAmount", LcAmount);
                                                command.Parameters.AddWithValue("@ExchangeRate", MyRow["ExchangeRate"].ToString());
                                                command.Parameters.AddWithValue("@ForeignCurrencyISOCode", MyRow["ForeignCurrencyISOCode"].ToString());
                                                command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                                command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                                command.CommandText = sqlMyEntry;
                                                command.ExecuteNonQuery();
                                            }
                                            command.Parameters.Clear();
                                        }

                                        //  decimal myAdminCharges = Common.toDecimal(model.LCCharges);
                                        //*** Cancellation Charges Transactions ***//

                                        #region Cancellation Charges Transactions
                                        //if (myAdminCharges > 0)
                                        //{
                                        //    string myChargesAccountCode = model.GeneralAccount;
                                        //    if (myChargesAccountCode != "")
                                        //    {
                                        //        #region credit charges account 
                                        //        string sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                        //        decimal ExchangeRate = 1;
                                        //        string ForeignCurrencyISOCode = "";
                                        //        decimal ForeignCurrencyAmount = 0;

                                        //        string sqlGLChartOfAccounts = "select CurrencyISOCode from  GLChartOfAccounts where AccountCode=@AccountCode";
                                        //        command.Transaction = transaction;
                                        //        command.Parameters.AddWithValue("@AccountCode", myChargesAccountCode);
                                        //        command.CommandText = sqlGLChartOfAccounts;
                                        //        if (command.ExecuteScalar() != null)
                                        //        {
                                        //            ForeignCurrencyISOCode = Common.toString(command.ExecuteScalar());
                                        //        }

                                        //        sqlGLChartOfAccounts = "select top 1 ExchangeRate from ExchangeRates where CurrencyISOCode=@CurrencyISOCode and AddedDate =@AddedDate";
                                        //        command.Transaction = transaction;
                                        //        command.Parameters.AddWithValue("@CurrencyISOCode", ForeignCurrencyISOCode);
                                        //        command.Parameters.AddWithValue("@AddedDate", Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00"));
                                        //        command.CommandText = sqlGLChartOfAccounts;
                                        //        if (command.ExecuteScalar() != null)
                                        //        {
                                        //            ExchangeRate = Common.toDecimal(command.ExecuteScalar());
                                        //        }
                                        //        command.Parameters.Clear();
                                        //        if (ForeignCurrencyISOCode != "" && ExchangeRate > 0)
                                        //        {
                                        //            if (!string.IsNullOrEmpty(model.FCCharges))
                                        //            {
                                        //                ForeignCurrencyAmount = Common.toDecimal(model.FCCharges);
                                        //            }
                                        //            else
                                        //            {
                                        //                ForeignCurrencyAmount = Common.toDecimal(ExchangeRate * myAdminCharges);
                                        //            }

                                        //            command.Parameters.Clear();
                                        //            ForeignCurrencyAmount = -1 * ForeignCurrencyAmount;
                                        //            command.Parameters.AddWithValue("@AccountCode", myChargesAccountCode);
                                        //            command.Parameters.AddWithValue("@GLTransactionTypeiCode", TypeiCode);
                                        //            command.Parameters.AddWithValue("@Memo", model.Description);
                                        //            command.Parameters.AddWithValue("@ForeignCurrencyAmount", ForeignCurrencyAmount);
                                        //            command.Parameters.AddWithValue("@LocalCurrencyAmount", -1 * myAdminCharges);
                                        //            command.Parameters.AddWithValue("@ExchangeRate", ExchangeRate);
                                        //            command.Parameters.AddWithValue("@ForeignCurrencyISOCode", ForeignCurrencyISOCode);
                                        //            command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                        //            command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                        //            command.CommandText = sqlMyEntry;
                                        //            command.ExecuteNonQuery();
                                        //            command.Parameters.Clear();
                                        //        }
                                        //        else
                                        //        {
                                        //            transaction.Rollback();
                                        //            isValid = false;
                                        //            model.ErrMessage = "Account currency not setup.";
                                        //            model.isPosted = false;
                                        //        }

                                        //        #endregion

                                        //        #region Agent debit account 
                                        //        if (isValid)
                                        //        {
                                        //            sqlMyEntry = "Insert into GLEntriesTemp (AccountCode, GLTransactionTypeiCode, Memo,ForeignCurrencyAmount,LocalCurrencyAmount,ExchangeRate,ForeignCurrencyISOCode, UserRefNo,TransactionNumber) Values (@AccountCode,@GLTransactionTypeiCode,@Memo,@ForeignCurrencyAmount,@LocalCurrencyAmount,@ExchangeRate,@ForeignCurrencyISOCode,@UserRefNo,@TransactionNumber)";
                                        //            ExchangeRate = 1;
                                        //            ForeignCurrencyISOCode = "";
                                        //            ForeignCurrencyAmount = 0;
                                        //            string myAgentAccountCode = "";
                                        //            sqlGLChartOfAccounts = "select CurrencyISOCode, AccountCode from  GLChartOfAccounts where prefix=@prefix";
                                        //            command.Transaction = transaction;
                                        //            command.Parameters.AddWithValue("@prefix", model.CustomerPrefix);
                                        //            command.CommandText = sqlGLChartOfAccounts;

                                        //            DataTable dtCoA = new DataTable();
                                        //            SqlDataAdapter DataAdaptter = new SqlDataAdapter(command);
                                        //            DataAdaptter.Fill(dtCoA);
                                        //            if (dtCoA != null)
                                        //            {
                                        //                if (dtCoA.Rows.Count > 0)
                                        //                {
                                        //                    ForeignCurrencyISOCode = Common.toString(dtCoA.Rows[0]["CurrencyISOCode"]);
                                        //                    myAgentAccountCode = Common.toString(dtCoA.Rows[0]["AccountCode"]);
                                        //                }
                                        //            }

                                        //            sqlGLChartOfAccounts = "select top 1 ExchangeRate from ExchangeRates where CurrencyISOCode=@CurrencyISOCode and AddedDate =@AddedDate";
                                        //            command.Transaction = transaction;
                                        //            command.Parameters.AddWithValue("@CurrencyISOCode", ForeignCurrencyISOCode);
                                        //            command.Parameters.AddWithValue("@AddedDate", Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00"));
                                        //            command.CommandText = sqlGLChartOfAccounts;
                                        //            if (command.ExecuteScalar() != null)
                                        //            {
                                        //                ExchangeRate = Common.toDecimal(command.ExecuteScalar());
                                        //            }
                                        //            command.Parameters.Clear();
                                        //            if (ForeignCurrencyISOCode != "" && ExchangeRate > 0 && myAgentAccountCode != "")
                                        //            {
                                        //                if (!string.IsNullOrEmpty(model.FCCharges))
                                        //                {
                                        //                    ForeignCurrencyAmount = Common.toDecimal(model.FCCharges);
                                        //                }
                                        //                else
                                        //                {
                                        //                    ForeignCurrencyAmount = Common.toDecimal(ExchangeRate * myAdminCharges);
                                        //                }
                                        //                command.Parameters.AddWithValue("@AccountCode", myAgentAccountCode);
                                        //                command.Parameters.AddWithValue("@GLTransactionTypeiCode", TypeiCode);
                                        //                command.Parameters.AddWithValue("@Memo", model.Description);
                                        //                command.Parameters.AddWithValue("@ForeignCurrencyAmount", ForeignCurrencyAmount);
                                        //                command.Parameters.AddWithValue("@LocalCurrencyAmount", myAdminCharges);
                                        //                command.Parameters.AddWithValue("@ExchangeRate", ExchangeRate);
                                        //                command.Parameters.AddWithValue("@ForeignCurrencyISOCode", ForeignCurrencyISOCode);
                                        //                command.Parameters.AddWithValue("@UserRefNo", cpUserRefNo);
                                        //                command.Parameters.AddWithValue("@TransactionNumber", model.PaymentNo);
                                        //                command.CommandText = sqlMyEntry;
                                        //                command.ExecuteNonQuery();
                                        //                command.Parameters.Clear();
                                        //            }
                                        //            else
                                        //            {
                                        //                transaction.Rollback();
                                        //                isValid = false;
                                        //                model.ErrMessage = "Agent account or currency not setup.";
                                        //                model.isPosted = false;
                                        //            }
                                        //        }
                                        //        #endregion
                                        //    }
                                        //}
                                        #endregion

                                        if (isValid)
                                        {
                                            transaction.Commit();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        transaction.Rollback();
                                        model.ErrMessage = Common.GetAlertMessage(1, "Error: " + ex.Message);
                                        model.isPosted = false;
                                    }
                                    finally
                                    {
                                        connection.Close();
                                    }
                                }
                                #endregion

                                #region Post for GL
                                if (isValid)
                                {
                                    DataTable lstGLEntriesTemp = DBUtils.GetDataTable("select * from GLEntriesTemp where UserRefNo='" + cpUserRefNo + "'");
                                    if (lstGLEntriesTemp != null)
                                    {
                                        if (lstGLEntriesTemp.Rows.Count > 0)
                                        {
                                            PostingViewModel PostingModel = PostingUtils.SaveGeneralJournelGLEntries(cpUserRefNo, "Transaction cancellation", "22CANCELENTRY", lstGLEntriesTemp, Profile.Id.ToString());
                                            if (PostingModel.isPosted)
                                            {
                                                DBUtils.ExecuteSQL("Update GLReferences set isPosted = '1' where ReferenceNo = '" + cpUserRefNo + "'");
                                                DBUtils.ExecuteSQL("Update CustomerTransactions set PaidGLReferenceID = '', status='Unpaid' where TransactionID = '" + model.TransactionID + "'");

                                                string TransactionID = "";
                                                string PaymentNo = "";
                                                string Status = "";
                                                string HoldAccountCode = "";
                                                string HoldGLReferenceID = "";
                                                string CustomerTransactionID = "";
                                                string BuyerName = "";
                                                //string BuyerRate = "";
                                                //string BuyerRateSC = "";
                                                string BuyerRateDC = "";
                                                string CustomerId = "";

                                                Guid guid1 = Guid.NewGuid();
                                                cpUserRefNo = guid1.ToString();

                                                string myConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings[Common.toString(SysPrefs.DefaultConnectionString)].ToString();
                                                using (SqlConnection connection = new SqlConnection(myConnectionString))
                                                {
                                                    SqlCommand command = connection.CreateCommand();
                                                    SqlTransaction transaction = null;

                                                    try
                                                    {
                                                        // BeginTransaction() Requires Open Connection
                                                        connection.Open();
                                                        transaction = connection.BeginTransaction();
                                                        string sqlVendors = "select * from Vendors where VendorID=@VendorID";
                                                        command.Transaction = transaction;
                                                        command.Parameters.AddWithValue("@VendorID", model.GeneralAccount);
                                                        command.CommandText = sqlVendors;
                                                        DataTable dtVendors = new DataTable();
                                                        SqlDataAdapter DataAdaptter1 = new SqlDataAdapter(command);
                                                        DataAdaptter1.Fill(dtVendors);
                                                        if (dtVendors != null)
                                                        {
                                                            if (dtVendors.Rows.Count > 0)
                                                            {
                                                                BuyerName = Common.toString(dtVendors.Rows[0]["VendorPrefix"]);
                                                            }
                                                        }
                                                        command.Parameters.Clear();

                                                        string CustomerName = "", Recipient = "", AgentName = "";
                                                        string sqlCustomerTransactions = "select AgentPrefix,SenderName,BuyerRateDC,Recipient,BuyerRate,BuyerRateSC,CustomerID,TransactionID, PaymentNo, Status, HoldAccountCode, HoldGLReferenceID, CustomerTransactionID from [CustomerTransactions] where PaymentNo=@PaymentNo";
                                                        command.Transaction = transaction;
                                                        command.Parameters.AddWithValue("@PaymentNo", model.PaymentNo);
                                                        // command.Parameters.AddWithValue("@Status", "Paid");
                                                        command.CommandText = sqlCustomerTransactions;
                                                        DataTable dtTransaction = new DataTable();
                                                        SqlDataAdapter DataAdaptter = new SqlDataAdapter(command);
                                                        DataAdaptter.Fill(dtTransaction);
                                                        if (dtTransaction != null)
                                                        {
                                                            if (dtTransaction.Rows.Count > 0)
                                                            {
                                                                TransactionID = Common.toString(dtTransaction.Rows[0]["TransactionID"]);
                                                                PaymentNo = Common.toString(dtTransaction.Rows[0]["PaymentNo"]);
                                                                Status = Common.toString(dtTransaction.Rows[0]["Status"]);
                                                                HoldAccountCode = Common.toString(dtTransaction.Rows[0]["HoldAccountCode"]);
                                                                HoldGLReferenceID = Common.toString(dtTransaction.Rows[0]["HoldGLReferenceID"]);
                                                                CustomerTransactionID = Common.toString(dtTransaction.Rows[0]["CustomerTransactionID"]);
                                                                CustomerId = Common.toString(dtTransaction.Rows[0]["CustomerID"]);
                                                                Recipient = Common.toString(dtTransaction.Rows[0]["Recipient"]);
                                                                CustomerName = Common.toString(dtTransaction.Rows[0]["SenderName"]);
                                                                BuyerRateDC = Common.toString(dtTransaction.Rows[0]["BuyerRateDC"]);
                                                                //BuyerRate = Common.toString(dtTransaction.Rows[0]["BuyerRate"]);
                                                                //BuyerRateSC = Common.toString(dtTransaction.Rows[0]["BuyerRateSC"]);
                                                                AgentName = Common.toString(dtTransaction.Rows[0]["AgentPrefix"]);
                                                            }
                                                        }
                                                        command.Parameters.Clear();

                                                        string sqlMyEntry = "Insert into FileEntriesTemp (CustomerId,CustomerName,TRANID,PaymentNo, Recipient,Buyer,BuyerRateSC,BuyerRateDC,BuyerRate, cpUserRefNo,AgentName,PostingDate) Values(@CustomerId, @CustomerName, @TRANID, @PaymentNo, @Recipient, @Buyer, @BuyerRateSC, @BuyerRateDC, @BuyerRate, @cpUserRefNo, @AgentName, @PostingDate)";
                                                        // command.Transaction = transaction;
                                                        //command.CommandTimeout = 60;
                                                        command.Parameters.AddWithValue("@CustomerId", CustomerId);
                                                        command.Parameters.AddWithValue("@CustomerName", CustomerName);
                                                        command.Parameters.AddWithValue("@TRANID", TransactionID);
                                                        command.Parameters.AddWithValue("@Recipient", Recipient);
                                                        command.Parameters.AddWithValue("@PaymentNo", PaymentNo);
                                                        command.Parameters.AddWithValue("@BuyerRateSC", model.BuyerRateSC);
                                                        command.Parameters.AddWithValue("@Buyer", BuyerName);
                                                        command.Parameters.AddWithValue("@BuyerRateDC", BuyerRateDC);
                                                        command.Parameters.AddWithValue("@BuyerRate", model.BuyerRate);
                                                        command.Parameters.AddWithValue("@cpUserRefNo", cpUserRefNo);
                                                        command.Parameters.AddWithValue("@AgentName", AgentName);
                                                        command.Parameters.AddWithValue("@PostingDate", Common.toDateTime(SysPrefs.PostingDate));
                                                        command.CommandText = sqlMyEntry;
                                                        command.ExecuteNonQuery();
                                                        command.Parameters.Clear();

                                                        string sqlTemp = "select * from FileEntriesTemp where cpUserRefNo=@VendorID";
                                                        command.Transaction = transaction;
                                                        command.Parameters.AddWithValue("@VendorID", cpUserRefNo);
                                                        command.CommandText = sqlTemp;
                                                        DataTable dtTempData = new DataTable();
                                                        SqlDataAdapter DataAdaptter11 = new SqlDataAdapter(command);
                                                        DataAdaptter11.Fill(dtTempData);
                                                        if (dtTempData != null)
                                                        {
                                                            if (dtTempData.Rows.Count > 0)
                                                            {
                                                                bool isMainValid = true;
                                                                ErrorMessages objErrorMessages = new ErrorMessages();
                                                                objErrorMessages = PostingUtils.SavePaidTransactionBridge("22PAIDENTRY", dtTempData, Profile.Id.ToString());
                                                                model.isPosted = false;
                                                                model.hideContent = false;
                                                                if (objErrorMessages != null)
                                                                {
                                                                    if (objErrorMessages.ErrorMessage.Count() > 0)
                                                                    {
                                                                        model.ErrMessage = Common.toString(objErrorMessages.ErrorMessage[0].Message);
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                transaction.Rollback();
                                                                model.ErrMessage = Common.GetAlertMessage(1, "Error : new entry does not created.");
                                                                model.isPosted = false;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            transaction.Rollback();
                                                            model.ErrMessage = Common.GetAlertMessage(1, "Error : new entry does not created.");
                                                            model.isPosted = false;
                                                        }
                                                        command.Parameters.Clear();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        model.ErrMessage = Common.GetAlertMessage(1, "Error :" + ex.Message.ToString());
                                                    }
                                                    finally
                                                    {
                                                        connection.Close();
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                model.ErrMessage = PostingModel.ErrMessage;
                                            }
                                        }
                                        else
                                        {
                                            model.ErrMessage = Common.GetAlertMessage(1, "No temp data found.. cp:" + cpUserRefNo);
                                        }
                                    }
                                    else
                                    {
                                        model.ErrMessage = Common.GetAlertMessage(1, "No temp data found. cp:" + cpUserRefNo);
                                    }
                                }
                                else
                                {
                                    model.ErrMessage = Common.GetAlertMessage(1, "No valid data");
                                }
                                #endregion
                            }
                            else
                            {
                                model.ErrMessage = Common.GetAlertMessage(1, "Data not found. id:" + model.cpGLReferenceID);
                            }
                        }
                        else
                        {
                            model.ErrMessage = Common.GetAlertMessage(1, "Invalid data");
                        }
                    }
                    else
                    {
                        model.CancellationMessage = Common.GetAlertMessage(1, "Transaction already cancelled.");
                    }
                }
                else
                {
                    model.ErrMessage = Common.GetAlertMessage(1, "Invalid Reference ID");
                }
                //}
                //else
                //{
                //    model.ErrMessage = Common.GetAlertMessage(1, "Incomplete data. PayoutCurrency not found.");
                //}
            }

            return View("~/Views/Postings/ReprocessTransactions.cshtml", model);
        }
        #endregion

        public ActionResult UploadFile()
        {
            ImportDepositFileViewModel model = new ImportDepositFileViewModel();
            string BaseFolder = @"C:\Users\STMDEV08\Documents\Clients"; //live path
                                                                        //foreach (string txtName in Directory.GetFiles(BaseFolder, "*.txt"))
                                                                        //{
                                                                        //HttpPostedFileBase file = Directory.GetFiles(BaseFolder, "*.txt");
                                                                        //}

            //    string filePath = Path.Combine(_siteRoot, @"C:\Users\STMDEV08\Documents\Clients.txt");
            //return File(filePath, "text/xml");
            //File file = new File("/path/to/uploaded/files", filename);
            return View("~/Views/Home/UploadFile.cshtml", model);
        }

        #region Extra functions
        //public ActionResult MenuPermission(MenuPermissionModel pMenu)
        //{
        //    UserPermissionViewModel checkViewModel = new UserPermissionViewModel();
        //    int kitchenId = Common.toInt(Session["CurrentKitchenId"]);
        //    //DateTime BussinessDate = Common.toDateTime(Session["BussinessDate"]);
        //    //checkViewModel.BussinessDate = BussinessDate;
        //    //checkViewModel.StartDate = BussinessDate;
        //    checkViewModel.UserID = "eb89021c-5fb5-4509-b9e7-9872bb636d9d";
        //    checkViewModel = getChecks(checkViewModel.UserID, 1);

        //    //pMenu.cpUserID = "eb89021c-5fb5-4509-b9e7-9872bb636d9d";
        //    //pMenu =LoadControls(pMenu.cpUserID, pMenu);
        //    //var pme = new UserModel();
        //    return View("~/Views/Users/administrator/AdministratorMenu.cshtml", pMenu);
        //    //return View();
        //}

        //private MenuPermissionModel LoadControls(string cpUserID, MenuPermissionModel pMenu)
        //{
        //    bool LeftOrRight = true;

        //    System.Web.UI.HtmlControls.HtmlGenericControl pnlLeft = new System.Web.UI.HtmlControls.HtmlGenericControl("DIV");
        //    System.Web.UI.HtmlControls.HtmlGenericControl pnlRight = new System.Web.UI.HtmlControls.HtmlGenericControl("DIV");
        //    if (!string.IsNullOrEmpty(cpUserID))
        //    {

        //        string Sqlgetuser = "Select UserID from [Users] where UserID='" + cpUserID + "'";
        //        DataTable getUserInfo = DBUtils.GetDataTable(Sqlgetuser);
        //        if (getUserInfo != null)
        //        {
        //            string SqlgetMenus = "Select MenuID,UserType,Title,URL,ParentMenuID from  DefaultMenus  Where UserType=99  And DefaultMenus.ParentMenuID=0 And Status=1";
        //            DataTable GetMainMenus = DBUtils.GetDataTable(SqlgetMenus);

        //            if (GetMainMenus != null && GetMainMenus.Rows.Count >= 1)
        //            {
        //                foreach (DataRow MainMenu in GetMainMenus.Rows)
        //                {
        //                    string SqlUserMenu = "Select MenuID from UsersMenu where MenuID=" + MainMenu["MenuID"].ToString() + " And UserID='" + cpUserID + "' ";

        //                    string usermenu = DBUtils.executeSqlGetSingle(SqlUserMenu);
        //                    if (string.IsNullOrEmpty(usermenu))
        //                    {
        //                        InseartMenu(MainMenu["MenuID"].ToString());
        //                    }

        //                    SyncUserMenus(Convert.ToInt32(MainMenu["MenuID"]));
        //                }
        //            }

        //            string SqlgetMenus1 = "Select um.UsersMenuID,um.UserID,um.MenuID,um.MenuResourceKey,um.CanEdit,um.CanAdd,um.CanDelete,um.Status,um.ActionControls,DefaultMenus.Title from  UsersMenu um,DefaultMenus Where um.UserID='" + cpUserID + "'  And um.MenuID= DefaultMenus.MenuID  And DefaultMenus.ParentMenuID=0";
        //            DataTable GetMenus1 = DBUtils.GetDataTable(SqlgetMenus1);

        //            if (GetMenus1 != null)
        //            {
        //                if (GetMenus1.Rows.Count > 0)
        //                {
        //                    //btnUpdate.Visible = true;
        //                    foreach (DataRow myMenuItem in GetMenus1.Rows)
        //                    {
        //                        bool CanShow = true;

        //                        if (CanShow)
        //                        {

        //                            System.Web.UI.HtmlControls.HtmlGenericControl divPanel = new System.Web.UI.HtmlControls.HtmlGenericControl("DIV");
        //                            divPanel.ID = "divPanel" + myMenuItem["UsersMenuID"].ToString();
        //                            divPanel.Attributes["class"] = "stm-pnl";
        //                            divPanel.Attributes["style"] = "float:left;";

        //                            System.Web.UI.HtmlControls.HtmlGenericControl divHeading = new System.Web.UI.HtmlControls.HtmlGenericControl("DIV");
        //                            divHeading.ID = "divHeading" + myMenuItem["UsersMenuID"].ToString();
        //                            divHeading.Attributes["class"] = "stm-pnl-heading";

        //                            CheckBox myCheckBox = new CheckBox();
        //                            myCheckBox.ID = "chk" + myMenuItem["UsersMenuID"].ToString();
        //                            myCheckBox.Text = myMenuItem["Title"].ToString();
        //                            if (myMenuItem["Status"].ToString() != null)
        //                            {
        //                                if (myMenuItem["Status"].ToString() == "True")
        //                                {
        //                                    myCheckBox.Checked = true;
        //                                }
        //                            }
        //                            divHeading.Controls.Add(myCheckBox);
        //                            divPanel.Controls.Add(divHeading);

        //                            System.Web.UI.HtmlControls.HtmlGenericControl divBody = new System.Web.UI.HtmlControls.HtmlGenericControl("DIV");
        //                            divBody.ID = "divBody" + myMenuItem["UsersMenuID"].ToString();
        //                            divBody.Attributes["class"] = "stm-pnl-bdy";
        //                            divBody.Style.Add("min-height", "300px");
        //                            GetSubMenus(Convert.ToInt32(myMenuItem["MenuID"].ToString()), ref divBody, pMenu);

        //                            //string SqlgetSubMenus = "Select um.UsersMenuID,um.UserID,um.MenuID,um.MenuResourceKey,um.CanEdit,um.CanAdd,um.CanDelete,um.Status,um.ActionControls,DefaultMenus.Title from  UsersMenu um,DefaultMenus Where um.UserID='" + cpUserID + "'  And um.MenuID= DefaultMenus.MenuID  And DefaultMenus.ParentMenuID=" + myMenuItem["MenuID"].ToString() + "";
        //                            //DataTable GetSubMenus = DBUtils.GetDataTable(SqlgetSubMenus);
        //                            //foreach (DataRow mySubMenuItem in GetSubMenus.Rows)
        //                            //{
        //                            //    System.Web.UI.HtmlControls.HtmlGenericControl divRow = new System.Web.UI.HtmlControls.HtmlGenericControl("DIV");
        //                            //    divRow.ID = "divRow" + mySubMenuItem["UsersMenuID"].ToString();
        //                            //    divRow.Attributes["class"] = "col-1-1";

        //                            //    System.Web.UI.HtmlControls.HtmlGenericControl divRowCell = new System.Web.UI.HtmlControls.HtmlGenericControl("DIV");
        //                            //    divRowCell.ID = "divRowCell" + mySubMenuItem["UsersMenuID"].ToString();
        //                            //    divRowCell.Attributes["class"] = "col-1-1";

        //                            //    CheckBox mySubCheckBox = new CheckBox();
        //                            //    mySubCheckBox.ID = "chk" + mySubMenuItem["UsersMenuID"].ToString();
        //                            //    mySubCheckBox.Text = mySubMenuItem["Title"].ToString();
        //                            //    mySubCheckBox.CssClass = "control-label";
        //                            //    if (mySubMenuItem["Status"].ToString() != null)
        //                            //    {
        //                            //        if (mySubMenuItem["Status"].ToString() == "True")
        //                            //        {
        //                            //            mySubCheckBox.Checked = true;
        //                            //        }
        //                            //    }

        //                            //    divRowCell.Controls.Add(mySubCheckBox);

        //                            //    LinkButton myHyperLink = new LinkButton();
        //                            //    myHyperLink.ID = "hl" + mySubMenuItem["UsersMenuID"].ToString();
        //                            //    myHyperLink.Text = "Settings";
        //                            //    myHyperLink.Attributes.Add("onclick", "OpenWin2('UserMenusSettings.aspx?UserID=" + cpUserID + "&MenuId=" + mySubMenuItem["MenuID"].ToString() + "&UserMenuId=" + mySubMenuItem["UsersMenuID"].ToString() + "');return false;");
        //                            //    if (mySubMenuItem["Status"].ToString() != null)
        //                            //    {
        //                            //        if (mySubMenuItem["Status"].ToString() != "True")
        //                            //        {
        //                            //            myHyperLink.Visible = false;
        //                            //        }
        //                            //    }


        //                            //    Label lblSpace = new Label();
        //                            //    lblSpace.Text = "&nbsp;&nbsp;";
        //                            //    divRowCell.Controls.Add(lblSpace);
        //                            //    divRowCell.Controls.Add(myHyperLink);
        //                            //    divRow.Controls.Add(divRowCell);
        //                            //    divBody.Controls.Add(UpdateSubMenu(cpUserID, Convert.ToInt32(mySubMenuItem["UsersMenuID"])));
        //                            //    divBody.Controls.Add(divRow);
        //                            //}

        //                            divPanel.Controls.Add(divBody);
        //                            if (LeftOrRight)
        //                            {
        //                                //pnlLeft.Controls.Add(divPanel);
        //                                StringBuilder generatedHtml = new StringBuilder();
        //                                using (var htmlStringWriter = new StringWriter(generatedHtml))
        //                                {
        //                                    using (var htmlTextWriter = new HtmlTextWriter(htmlStringWriter))
        //                                    {
        //                                        divPanel.RenderControl(htmlTextWriter);
        //                                        pMenu.PnlLeft += generatedHtml.ToString();
        //                                    }
        //                                }
        //                                LeftOrRight = false;
        //                            }
        //                            else
        //                            {
        //                                //pnlRight.Controls.Add(divPanel);
        //                                StringBuilder generatedHtml = new StringBuilder();
        //                                using (var htmlStringWriter = new StringWriter(generatedHtml))
        //                                {
        //                                    using (var htmlTextWriter = new HtmlTextWriter(htmlStringWriter))
        //                                    {
        //                                        divPanel.RenderControl(htmlTextWriter);
        //                                        pMenu.PnlRight += generatedHtml.ToString();
        //                                    }
        //                                }
        //                                LeftOrRight = true;
        //                            }
        //                        }
        //                    }
        //                    //StringBuilder generatedHtmlpnlLeft = new StringBuilder();
        //                    //using (var htmlStringWriter = new StringWriter(generatedHtmlpnlLeft))
        //                    //{
        //                    //    using (var htmlTextWriter = new HtmlTextWriter(htmlStringWriter))
        //                    //    {
        //                    //        pnlLeft.RenderControl(htmlTextWriter);
        //                    //        DynamicHtml = pnlLeft.ToString();
        //                    //    }
        //                    //}



        //                    //pMenu.table = DynamicHtml;
        //                }
        //                else
        //                {
        //                    //btnUpdate.Visible = true;
        //                }
        //            }
        //        }
        //    }
        //    return pMenu;
        //}

        //private void SyncUserMenus(int pParentID)
        //{
        //    string SqlgetMenusChild = "Select MenuID,UserType,Title,URL,ParentMenuID from  DefaultMenus  Where UserType=99  And DefaultMenus.ParentMenuID=" + pParentID + "";
        //    DataTable GetMainMenusChild = DBUtils.GetDataTable(SqlgetMenusChild);
        //    if (GetMainMenusChild != null && GetMainMenusChild.Rows.Count >= 1)
        //    {
        //        foreach (DataRow MainmenuChild in GetMainMenusChild.Rows)
        //        {
        //            string SqlUserMenuChild = "Select MenuID from UsersMenu where MenuID=" + MainmenuChild["MenuID"].ToString() + " And UserID=99";//'" + cpUserID + "' ";
        //            string usermenuChild = DBUtils.executeSqlGetSingle(SqlUserMenuChild);
        //            if (string.IsNullOrEmpty(usermenuChild))
        //            {
        //                InseartMenu(MainmenuChild["MenuID"].ToString());
        //            }
        //            SyncUserMenus(Convert.ToInt32(MainmenuChild["MenuID"]));
        //        }
        //    }
        //}
        //private void GetSubMenus(int pMenuID, ref System.Web.UI.HtmlControls.HtmlGenericControl pDivBody, MenuPermissionModel pMenu)
        //{
        //    string sqlSubMenus = "Select um.UsersMenuID,um.UserID,um.MenuID,um.MenuResourceKey,um.CanEdit,um.CanAdd,um.CanDelete,um.Status,um.ActionControls,DefaultMenus.Title from  UsersMenu um,DefaultMenus Where um.UserID='" + pMenu.cpUserID + "' And um.MenuID= DefaultMenus.MenuID  And DefaultMenus.ParentMenuID=" + pMenuID.ToString() + "";
        //    DataTable dtSubMenus = DBUtils.GetDataTable(sqlSubMenus);
        //    foreach (DataRow mySubMenuItem in dtSubMenus.Rows)
        //    {
        //        System.Web.UI.HtmlControls.HtmlGenericControl divRow = new System.Web.UI.HtmlControls.HtmlGenericControl("DIV");
        //        divRow.ID = "divRow" + mySubMenuItem["UsersMenuID"].ToString();
        //        divRow.Attributes["class"] = "col-1-1";

        //        System.Web.UI.HtmlControls.HtmlGenericControl divRowCell = new System.Web.UI.HtmlControls.HtmlGenericControl("DIV");
        //        divRowCell.ID = "divRowCell" + mySubMenuItem["UsersMenuID"].ToString();
        //        divRowCell.Attributes["class"] = "col-1-1";


        //        CheckBox mySubCheckBox = new CheckBox();
        //        mySubCheckBox.ID = "chk" + mySubMenuItem["UsersMenuID"].ToString();
        //        mySubCheckBox.Text = mySubMenuItem["Title"].ToString();
        //        mySubCheckBox.CssClass = "control-label";
        //        if (mySubMenuItem["Status"].ToString() != null)
        //        {
        //            if (mySubMenuItem["Status"].ToString() == "True")
        //            {
        //                mySubCheckBox.Checked = true;
        //            }
        //        }

        //        divRowCell.Controls.Add(mySubCheckBox);

        //        LinkButton myHyperLink = new LinkButton();
        //        myHyperLink.ID = "hl" + mySubMenuItem["UsersMenuID"].ToString();
        //        myHyperLink.Text = "Settings";
        //        myHyperLink.Attributes.Add("href", "#");
        //        myHyperLink.Attributes.Add("onclick", "OpenWin2('UserMenusSettings.aspx?UserID=" + pMenu.cpUserID + "&MenuId=" + mySubMenuItem["MenuID"].ToString() + "&UserMenuId=" + mySubMenuItem["UsersMenuID"].ToString() + "');return false;");
        //        if (mySubMenuItem["Status"].ToString() != null)
        //        {
        //            if (mySubMenuItem["Status"].ToString() != "True")
        //            {
        //                myHyperLink.Visible = false;
        //            }
        //        }

        //        Label lblSpace = new Label();
        //        lblSpace.Text = "&nbsp;&nbsp;";
        //        divRowCell.Controls.Add(lblSpace);
        //        divRowCell.Controls.Add(myHyperLink);
        //        divRow.Controls.Add(divRowCell);

        //        pDivBody.Controls.Add(divRow);
        //        GetSubMenus(Convert.ToInt32(mySubMenuItem["MenuID"].ToString()), ref pDivBody, pMenu);
        //    }
        //    //return divRow;
        //}
        //private void SaveControls(MenuPermissionModel pMenu)
        //{
        //    bool LeftOrRight = true;
        //    string Sqlgetuser = "Select UserID from [Users] where UserID='" + pMenu.cpUserID + "'";
        //    DataTable getUserInfo = DBUtils.GetDataTable(Sqlgetuser);
        //    string SqlUpdateUserMenu = "UPDATE [dbo].[UsersMenu] set [Status] = @Status, ActionControls=@ActionControls  WHERE UsersMenuID=@UsersMenuID";
        //    string UpdateActionControls = "";
        //    SqlTransaction Trans = null;

        //    using (SqlConnection con = Common.getConnection())
        //    {
        //        SqlCommand Updatecommand = con.CreateCommand();
        //        Trans = con.BeginTransaction();
        //        if (getUserInfo != null)
        //        {
        //            string SqlgetMenus = "Select um.UsersMenuID,um.UserID,um.MenuID,um.MenuResourceKey,um.CanEdit,um.CanAdd,um.CanDelete,um.Status,um.ActionControls,DefaultMenus.Title from  UsersMenu um,DefaultMenus Where um.UserID='" + pMenu.cpUserID + "' And um.MenuID= DefaultMenus.MenuID  And DefaultMenus.ParentMenuID=0";
        //            DataTable GetMenus = DBUtils.GetDataTable(SqlgetMenus);
        //            try
        //            {
        //                if (GetMenus != null)
        //                {
        //                    foreach (DataRow myMenuItem in GetMenus.Rows)
        //                    {
        //                        bool CanShow = true;
        //                        if (CanShow)
        //                        {
        //                            string CName = "chk" + myMenuItem["UsersMenuID"].ToString();
        //                            if (LeftOrRight)
        //                            {
        //                                //string checkbox=pMenu.PnlLeft.Substring(pMenu.PnlLeft.IndexOf(CName),)
        //                                CheckBox chkBox = new CheckBox();//(CheckBox)pnlLeft.FindControl(CName);
        //                                Updatecommand.Parameters.AddWithValue("ActionControls", "");
        //                                Updatecommand.Parameters.AddWithValue("Status", chkBox.Checked);
        //                            }
        //                            else
        //                            {
        //                                CheckBox chkBox = new CheckBox();//(CheckBox)pnlRight.FindControl(CName);
        //                                Updatecommand.Parameters.AddWithValue("ActionControls", "");
        //                                Updatecommand.Parameters.AddWithValue("Status", chkBox.Checked);
        //                            }

        //                            Updatecommand.Parameters.AddWithValue("UsersMenuID", myMenuItem["UsersMenuID"].ToString());
        //                            Updatecommand.Transaction = Trans;
        //                            Updatecommand.CommandText = SqlUpdateUserMenu;
        //                            Updatecommand.ExecuteNonQuery();
        //                            Updatecommand.Parameters.Clear();

        //                            string SqlgetSubMenus = "Select um.UsersMenuID,um.UserID,um.MenuID,um.MenuResourceKey,um.CanEdit,um.CanAdd,um.CanDelete,um.Status,um.ActionControls,DefaultMenus.Title from  UsersMenu um,DefaultMenus Where um.UserID=@UserID  And um.MenuID= DefaultMenus.MenuID  And DefaultMenus.ParentMenuID=@ParentMenuID";
        //                            SqlDataAdapter getSMuenus = new SqlDataAdapter();

        //                            Updatecommand.Parameters.AddWithValue("@UserID", pMenu.cpUserID);
        //                            Updatecommand.Parameters.AddWithValue("@ParentMenuID", myMenuItem["MenuID"].ToString());
        //                            Updatecommand.CommandText = SqlgetSubMenus;
        //                            Updatecommand.Transaction = Trans;
        //                            getSMuenus.SelectCommand = Updatecommand;
        //                            DataTable GetSubMenus = new DataTable();
        //                            getSMuenus.Fill(GetSubMenus);
        //                            Updatecommand.Parameters.Clear();

        //                            foreach (DataRow mySubMenuItem in GetSubMenus.Rows)
        //                            {
        //                                string myActionControls = "";
        //                                string CSubName = "chk" + mySubMenuItem["UsersMenuID"].ToString();
        //                                if (LeftOrRight)
        //                                {
        //                                    CheckBox mySubCheckBox = new CheckBox();// (CheckBox)pnlLeft.FindControl(CSubName);
        //                                    if (mySubCheckBox.Checked)
        //                                    {
        //                                        if (mySubMenuItem["Status"].ToString() == "False")
        //                                        {
        //                                            myActionControls = Common.GetActionControls(Convert.ToInt32(mySubMenuItem["MenuID"].ToString()));
        //                                        }
        //                                        else
        //                                        {
        //                                            myActionControls = mySubMenuItem["ActionControls"].ToString();
        //                                        }
        //                                    }
        //                                    Updatecommand.Parameters.AddWithValue("@ActionControls", myActionControls);
        //                                    Updatecommand.Parameters.AddWithValue("Status", mySubCheckBox.Checked);

        //                                }
        //                                else
        //                                {
        //                                    CheckBox mySubCheckBox = new CheckBox();//(CheckBox)pnlRight.FindControl(CSubName);
        //                                    if (mySubCheckBox.Checked)
        //                                    {
        //                                        if (mySubMenuItem["Status"].ToString() == "False")
        //                                        {
        //                                            myActionControls = Common.GetActionControls(Convert.ToInt32(mySubMenuItem["MenuID"].ToString()));
        //                                        }
        //                                        else
        //                                        {
        //                                            myActionControls = mySubMenuItem["ActionControls"].ToString();
        //                                        }
        //                                    }

        //                                    Updatecommand.Parameters.AddWithValue("@ActionControls", myActionControls);
        //                                    Updatecommand.Parameters.AddWithValue("Status", mySubCheckBox.Checked);
        //                                }

        //                                Updatecommand.Parameters.AddWithValue("UsersMenuID", mySubMenuItem["UsersMenuID"].ToString());
        //                                Updatecommand.Transaction = Trans;
        //                                Updatecommand.CommandText = SqlUpdateUserMenu;
        //                                Updatecommand.ExecuteNonQuery();
        //                                Updatecommand.Parameters.Clear();
        //                            }
        //                            if (LeftOrRight)
        //                            {
        //                                LeftOrRight = false;
        //                            }
        //                            else
        //                            {
        //                                LeftOrRight = true;
        //                            }
        //                        }
        //                    }
        //                }

        //                Trans.Commit();

        //                //phMessage.Visible = true;
        //                //lblMessage.Text = iRemitifyAccounts.Core.Common.GetAlertMessage(0, "Menues Updated successfully");
        //            }
        //            catch (Exception msg)
        //            {
        //                Trans.Rollback();
        //                //phMessage.Visible = true;
        //                //lblMessage.Text = msg.Message;

        //            }
        //        }
        //    }
        //}
        //[HttpPost]
        //[ValidateInput(false)]
        //public void UpdateMenu(MenuPermissionModel pMenu, HttpPostedFileBase file)
        //{
        //    if (pMenu.SecurityStamp != "")
        //    {
        //        if (!(Common.ValidatePinCode(pMenu.cpUserID, pMenu.SecurityStamp)))
        //        {
        //            //phMessage.Visible = true;
        //            //lblMessage.Text = iRemitifyAccounts.Core.Common.GetAlertMessage(1, "Invalid Pin code entered");
        //        }
        //        else
        //        {
        //            SaveControls(pMenu);
        //        }

        //    }
        //    else
        //    {
        //        //phMessage.Visible = true;
        //        //lblMessage.Text = iRemitifyAccounts.Core.Common.GetAlertMessage(1, "Please enter your pin code ");
        //    }
        //}
        //private void InseartMenu(string MenuId)
        //{
        //    string SqlInsertUserMenu = "INSERT INTO [dbo].[UsersMenu]  ([UserID],[MenuID],[MenuResourceKey],[CanEdit],[CanAdd],[CanDelete] ,[Status],[ActionControls])";
        //    SqlInsertUserMenu += "VALUES (@UserID,@MenuID,@MenuResourceKey,@CanEdit,@CanAdd,@CanDelete,@Status,@ActionControls) ";

        //    using (SqlConnection con = Common.getConnection())
        //    {
        //        SqlCommand InsertUserMenuCommand = con.CreateCommand();

        //        InsertUserMenuCommand.Parameters.AddWithValue("@UserID", 99);
        //        InsertUserMenuCommand.Parameters.AddWithValue("@MenuID", MenuId);
        //        InsertUserMenuCommand.Parameters.AddWithValue("@MenuResourceKey", MenuId);
        //        InsertUserMenuCommand.Parameters.AddWithValue("@CanEdit", 0);
        //        InsertUserMenuCommand.Parameters.AddWithValue("@CanAdd", 0);
        //        InsertUserMenuCommand.Parameters.AddWithValue("@CanDelete", 0);
        //        InsertUserMenuCommand.Parameters.AddWithValue("@Status", 0);
        //        InsertUserMenuCommand.Parameters.AddWithValue("@ActionControls", "");

        //        InsertUserMenuCommand.CommandText = SqlInsertUserMenu;
        //        InsertUserMenuCommand.ExecuteNonQuery();
        //        InsertUserMenuCommand.Parameters.Clear();
        //    }
        //}

        ////public static IEnumerable<CheckListViewModel> GetChecksList(KitchenChecksRequest pKitchenChecksRequest)
        ////{
        ////    IEnumerable<CheckListViewModel> pcheckListViewModel = new List<CheckListViewModel>();
        ////    if (pKitchenChecksRequest.KitchenId > 0)
        ////    {
        ////        string sqlCheckList = "SELECT * from CheckList where KitchenId='" + pKitchenChecksRequest.KitchenId + "' and CheckTypeId='" + pKitchenChecksRequest.CheckTypeId + "'";
        ////        DataTable dtCheck = DBUtils.GetDataTable(sqlCheckList);
        ////        if (dtCheck != null && dtCheck.Rows.Count > 0)
        ////        {
        ////            foreach (DataRow myDataRow in dtCheck.Rows)
        ////            {
        ////                CheckListViewModel myChecks = new CheckListViewModel();
        ////                myChecks.CheckListId = Common.toInt(myDataRow["CheckListId"]);
        ////                myChecks.CheckText = myDataRow["CheckText"].ToString();
        ////                myChecks.Status = Common.toBool(myDataRow["Status"]);
        ////                pcheckListViewModel = (new[] { myChecks }).Concat(pcheckListViewModel);
        ////            }
        ////        }
        ////    }
        ////    return pcheckListViewModel;
        ////}


        //public static UserPermissionViewModel getChecks(string pKitchenId, int pCheckTypeId)
        //{
        //    //KitchenHealthEntities db = new KitchenHealthEntities();
        //    iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
        //    UserPermissionViewModel checkViewModel = new UserPermissionViewModel();

        //    string CheckTypeName = "";
        //    //DateTime StartDay = BussinessDate, EndDay = BussinessDate;
        //    if (pCheckTypeId == 1)
        //    {
        //        CheckTypeName = "Opening Checks";
        //    }
        //    else if (pCheckTypeId == 2)
        //    {
        //        CheckTypeName = "Closing Checks";
        //    }
        //    //else if (pCheckTypeId == 3)
        //    //{
        //    //    DayOfWeek day = BussinessDate.DayOfWeek;
        //    //    //int days = day - DayOfWeek.Monday;
        //    //    int Cday = 0;
        //    //    if (day == DayOfWeek.Sunday)
        //    //    {
        //    //        Cday = 6;
        //    //    }
        //    //    else
        //    //    {
        //    //        Cday = day - DayOfWeek.Monday;
        //    //    }
        //    //    StartDay = Common.toDateTime(BussinessDate.AddDays(-Cday).ToShortDateString());
        //    //    EndDay = StartDay.AddDays(6);
        //    //    CheckTypeName = "Weekly Checks    From Date:  " + StartDay.ToString("dd/M/yyyy") + "     To Date:  " + EndDay.ToString("dd/M/yyyy");
        //    //}
        //    //else if (pCheckTypeId == 4)
        //    //{

        //    //    DateTime now = BussinessDate;
        //    //    StartDay = new DateTime(now.Year, now.Month, 1);
        //    //    EndDay = StartDay.AddMonths(1).AddDays(-1);
        //    //    //  DateTime Month = now.Month.ToString("MMM");
        //    //    int Year = now.Year;
        //    //    //  string month = now.ToString("MMM", CultureInfo.InvariantCulture);
        //    //    CheckTypeName = "Monthly Checks:     " + now.ToString("MMMM", CultureInfo.InvariantCulture) + "," + Year;
        //    //    //  CheckTypeName = "Monthly Checks:";
        //    //}
        //    //checkViewModel.BussinessDate = BussinessDate;
        //    //checkViewModel.StartDate = StartDay;
        //    //checkViewModel.EndDate = EndDay;
        //    var kitchenCheckList = db.DefaultMenus.Where(x => x.ParentMenuID == 0).ToList();
        //    //if (pCheckTypeId == 1 || pCheckTypeId == 2)
        //    //{
        //    //    kitchenCheckList = db.KitchenCheckList.Where(x => x.KitchenId == pKitchenId && x.AddedDate == BussinessDate && x.CheckTypeId == pCheckTypeId).Take(1).SingleOrDefault();
        //    //}
        //    //else if (pCheckTypeId == 3 || pCheckTypeId == 4)
        //    //{
        //    //    kitchenCheckList = db.KitchenCheckList.Where(x => x.KitchenId == pKitchenId && x.StartDate == StartDay && x.EndDate == EndDay && x.CheckTypeId == pCheckTypeId).Take(1).SingleOrDefault();
        //    //}
        //    if (kitchenCheckList != null)
        //    {
        //        //UserMenu UI = new Models.UserMenu();
        //        MenuItemViewModel MI = new MenuItemViewModel();
        //        IEnumerable<MenuItemViewModel> ListMenu = new List<MenuItemViewModel>();
        //        IEnumerable<UserMenuViewModel> UMVM = new List<UserMenuViewModel>();
        //        foreach (var Menu in kitchenCheckList)
        //        {
        //            MI.MenuID = Menu.MenuID;
        //            MI.Title = Menu.Title;
        //            MI.subMenu = GetChecksList(MI.MenuID);
        //            ListMenu = (new[] { MI }).Concat(ListMenu);
        //        }
        //    }

        //    return checkViewModel;
        //}

        //public static IEnumerable<UserMenuViewModel> GetChecksList(int MenuID)
        //{
        //    IEnumerable<UserMenuViewModel> UMVM = new List<UserMenuViewModel>();
        //    if (MenuID > 0)
        //    {
        //        UserMenuViewModel UsrMVM = new UserMenuViewModel();
        //        string sqlCheckList = "SELECT UM.MenuID,UM.UserID,DM.ParentMenuID, DM.Title from UsersMenu UM inner join DefaultMenus DM on UM.MenuID=DM.MenuID where MenuId=" + MenuID;
        //        DataTable dtCheck = DBUtils.GetDataTable(sqlCheckList);
        //        if (dtCheck != null && dtCheck.Rows.Count > 0)
        //        {
        //            foreach (DataRow myDataRow in dtCheck.Rows)
        //            {
        //                UsrMVM.MenuID = Common.toInt(myDataRow["MenuID"]);
        //                UsrMVM.ParentMenuID = Common.toInt(myDataRow["ParentMenuID"]);
        //                UsrMVM.Title = Common.toString(myDataRow["Title"]);
        //                UsrMVM.UserID = Common.toString(myDataRow["UserID"]);

        //                UMVM = (new[] { UsrMVM }).Concat(UMVM);
        //            }
        //        }
        //    }
        //    return UMVM;
        //}
        #endregion

    }
}