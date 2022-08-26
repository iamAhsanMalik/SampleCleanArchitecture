using Dapper;
using DHAAccounts.Models;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Mvc;


namespace Accounting.Controllers
{
    public class CurrencyController : Controller
    {
        ////private iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
        ApplicationUser profile = ApplicationUser.GetUserProfile();
        private AppDbContext _dbcontext = new AppDbContext();
        //// private AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());
        // GET: Currency

        #region Currency Management
        public ActionResult Index()
        {
            //// DBUtils.executeSqlGetSingle("delete ListCurrencies where CurrencyISOCode='EUt'");
            return View("~/Views/Currency/Index.cshtml");
        }


        public ActionResult CheckList_Read([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(5) FROM ListCurrencies /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                          CurrencyISOCode,CurrencyName ,CurrencySymbol, (select top 1 CountryName from ListCountries where ListCountries.CountryISONumericCode=ListCurrencies.CountryISOCode ) as  CountryISOCode, CurrencyStatus, MinRateLimit, MaxRateLimit, CurrenyType, FORMAT ( LastUpdatedDate, 'd', 'en-gb' ) as LastUpdatedDate, CurrencyRate FROM ListCurrencies
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";
            //     const string selectQuery = @"SELECT
            //CheckListId, CheckTypeId, (select CheckType  from CheckTypes where CheckTypes.CheckTypeId = CheckList.CheckTypeId ) as CheckTypeName, CheckText, case when Status = 'true' then 'Active' else 'Inactive' end as StatusName, Status FROM CheckList";
            SqlBuilder builder = new SqlBuilder();
            var count = builder.AddTemplate(countQuery);

            int CurrentPage = 0; int CurrentPageSize = 25;
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
            //else
            //{
            //    builder.Where("CurrencyStatus=1");
            //}

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("CurrencyISOCode");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);

            var rows = _dbcontext.Query<CurrencyModel>(selector.RawSql, selector.Parameters);

            //  rows.Each(item => Console.WriteLine($"({ item.AccountCode}): { item.AccountName}
            //{                GetParentsString(items, item)}"));            
            //}
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }
        public ActionResult Create()
        {
            return View("~/Views/Currency/Create.cshtml");
        }

        [HttpPost]
        public ActionResult Create(CurrencyModel pCurrencyModel)
        {
            if (ModelState.IsValid)
            {
                string eCurrencyISOCode = DBUtils.executeSqlGetSingle("select CurrencyISOCode from   ListCurrencies Where CurrencyISOCode='" + pCurrencyModel.CurrencyISOCode + "'");
                if (Common.toString(eCurrencyISOCode).Trim() == "")
                {
                    string SqlInsert = "Insert Into ListCurrencies (AutoCreateUnPaidAccount,AutoCreatUnEarnedAdminAccount,AutoCreateUnEarnedAgentAccount,CurrencyISOCode,CurrencySymbol,CurrencyName,MinRateLimit,MaxRateLimit,CurrenyType,isProfit,CountryISOCode,CurrencyStatus,RevalueYN,DateAdded)";
                    SqlInsert += "Values(@AutoCreateUnPaidAccount,@AutoCreatUnEarnedAdminAccount,@AutoCreateUnEarnedAgentAccount,@CurrencyISOCode, @CurrencySymbol, @CurrencyName, @MinRateLimit, @MaxRateLimit, @CurrenyType, @isProfit,@CountryISOCode,@CurrencyStatus,@RevalueYN,@DateAdded)";
                    SqlConnection conn = Common.getConnection();
                    ////using (SqlConnection conn = Common.getConnection())
                    ////{
                    SqlCommand insert = conn.CreateCommand();
                    try
                    {
                        insert.Parameters.AddWithValue("@CurrencyISOCode", pCurrencyModel.CurrencyISOCode);
                        insert.Parameters.AddWithValue("@CurrencySymbol", pCurrencyModel.CurrencySymbol);
                        insert.Parameters.AddWithValue("@CurrencyName", pCurrencyModel.CurrencyName);
                        //insert.Parameters.AddWithValue("@LockYesNo", pCurrencyModel.LockYesNo);
                        insert.Parameters.AddWithValue("@MinRateLimit", pCurrencyModel.MinRateLimit);
                        insert.Parameters.AddWithValue("@MaxRateLimit", pCurrencyModel.MaxRateLimit);
                        insert.Parameters.AddWithValue("@CurrenyType", pCurrencyModel.CurrenyType);
                        insert.Parameters.AddWithValue("@isProfit", pCurrencyModel.isProfit);
                        insert.Parameters.AddWithValue("@CountryISOCode", pCurrencyModel.CountryISOCode);
                        insert.Parameters.AddWithValue("@CurrencyStatus", pCurrencyModel.CurrencyStatus);
                        insert.Parameters.AddWithValue("@RevalueYN", pCurrencyModel.RevalueYN);
                        insert.Parameters.AddWithValue("@AutoCreateUnPaidAccount", pCurrencyModel.CreateHoldAccount);
                        insert.Parameters.AddWithValue("@AutoCreatUnEarnedAdminAccount", pCurrencyModel.CreateAdminFeeAccount);
                        insert.Parameters.AddWithValue("@AutoCreateUnEarnedAgentAccount", pCurrencyModel.CreateAgentFeeAccount);
                        insert.Parameters.AddWithValue("@DateAdded", DateTime.Now);
                        insert.CommandText = SqlInsert;
                        int RowsAffected = insert.ExecuteNonQuery();
                        if (RowsAffected == 1)
                        {
                            //string HoldAccount = "";
                            //string AdminChargesAccount = "";
                            //string AgentChargesAccount = "";
                            //if (pCurrencyModel.CreateHoldAccount)
                            //{
                            //    AddAccountCodeResponse response = Common.AddGLAccountCode(SysPrefs.TransactionHoldingAccount, pCurrencyModel.CurrencyISOCode, pCurrencyModel.CurrencyISOCode);
                            //    if (response.isAdded)
                            //    {
                            //        HoldAccount = response.AccountCode;
                            //    }
                            //}

                            //if (pCurrencyModel.CreateAdminFeeAccount)
                            //{
                            //    AddAccountCodeResponse response = Common.AddGLAccountCode(SysPrefs.UnEarnedServiceFeeAccount, pCurrencyModel.CurrencyISOCode + "-UNEARNED ADMIN CHARGES INCOME", pCurrencyModel.CurrencyISOCode);
                            //    if (response.isAdded)
                            //    {
                            //        AdminChargesAccount = response.AccountCode;
                            //    }
                            //}

                            //if (pCurrencyModel.CreateAgentFeeAccount)
                            //{
                            //    AddAccountCodeResponse response = Common.AddGLAccountCode(SysPrefs.UnEarnedServiceFeeAccount, pCurrencyModel.CurrencyISOCode + "-UNEARNED AGENT CHARGES INCOME", pCurrencyModel.CurrencyISOCode);
                            //    if (response.isAdded)
                            //    {
                            //        AgentChargesAccount = response.AccountCode;
                            //    }
                            //}

                            //DBUtils.ExecuteSQL("Update ListCurrencies Set HoldAccountCode='" + HoldAccount + "', UnEarnedAdminAccountCode='" + AdminChargesAccount + "',  UnEarnedAgentAccountCode='" + AgentChargesAccount + "' where CurrencyISOCode='" + pCurrencyModel.CurrencyISOCode + "'");

                            TempData["Success"] = "Currency successfully added";
                            return View("~/Views/Currency/Create.cshtml", pCurrencyModel);
                        }
                        else
                        {
                            TempData["Failed"] = "Error occured: ";
                            return View("~/Views/Currency/Create.cshtml", pCurrencyModel);
                        }
                    }
                    catch (Exception ex)
                    {
                        TempData["Failed"] = "Error occured: " + ex.Message.ToString();
                        return View("~/Views/Currency/Create.cshtml", pCurrencyModel);
                    }
                    ////  }
                    conn.Close();
                }
                else
                {
                    TempData["Failed"] = "Error occured: Currency ISO code already exists";
                    return View("~/Views/Currency/Create.cshtml", pCurrencyModel);
                }
            }
            else
            {
                TempData["Failed"] = "Data missing. Please check all fields and try again";
                return View("~/Views/Currency/Create.cshtml", pCurrencyModel);
            }
        }

        //[AcceptVerbs(HttpVerbs.Post)]
        //public ActionResult EditingCustom_Destroy([DataSourceRequest] DataSourceRequest request,
        //   [Bind(Prefix = "models")]IEnumerable<CurrencyModel> products)
        //{
        //    foreach (var product in products)
        //    {

        //       DataSource.Destroy(product);
        //    }

        //    return Json(products.ToDataSourceResult(request, ModelState));
        //}
        //[AcceptVerbs(HttpVerbs.Post)]
        //public ActionResult EditingCustom_Update([DataSourceRequest] DataSourceRequest request,
        //  [Bind(Prefix = "models")]IEnumerable<CurrencyModel> products)
        //{
        //    if (products != null && ModelState.IsValid)
        //    {
        //        foreach (var product in products)
        //        {
        //            ListCurrencies.Update(product);
        //        }
        //    }

        //    return Json(products.ToDataSourceResult(request, ModelState));
        //}

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult EditingInline_Update([DataSourceRequest] DataSourceRequest request, CurrencyModel product)
        {
            if (product != null)
            {
                EditingCustom_Update(product);
            }

            return Json(new[] { product }.ToDataSourceResult(request, ModelState));
        }

        public void EditingCustom_Update(CurrencyModel pCurrencyModel)
        {
            if (pCurrencyModel.CurrencyISOCode != "" && pCurrencyModel.MinRateLimit > 0.0001m && pCurrencyModel.MinRateLimit > 0.0001m)
            {
                string UpdateSQL = "Update ListCurrencies Set ";
                UpdateSQL += "MinRateLimit=@MinRateLimit,MaxRateLimit=@MaxRateLimit,CurrencyStatus=@CurrencyStatus,LastUpdatedDate=@LastUpdatedDate,LastUpdatedBy=@LastUpdatedBy Where CurrencyISOCode = @CurrencyISOCode";
                SqlConnection con = Common.getConnection();
                //////using (SqlConnection con = Common.getConnection())
                //////{
                SqlCommand updateCommand = con.CreateCommand();
                try
                {
                    if (pCurrencyModel.MaxRateLimit > pCurrencyModel.MinRateLimit)
                    {
                        //updateCommand.Parameters.AddWithValue("@UnEarnedAgentAccountCode", pCurrencyModel.UnEarnedAgentAccountCode);
                        //updateCommand.Parameters.AddWithValue("@UnEarnedAdminAccountCode", pCurrencyModel.UnEarnedAdminAccountCode);
                        //updateCommand.Parameters.AddWithValue("@CurrencyRate", pCurrencyModel.CurrencyRate);
                        updateCommand.Parameters.AddWithValue("@MinRateLimit", pCurrencyModel.MinRateLimit);
                        updateCommand.Parameters.AddWithValue("@MaxRateLimit", pCurrencyModel.MaxRateLimit);
                        updateCommand.Parameters.AddWithValue("@CurrencyStatus", pCurrencyModel.CurrencyStatus);
                        updateCommand.Parameters.AddWithValue("@LastUpdatedDate", DateTime.Now);
                        updateCommand.Parameters.AddWithValue("@LastUpdatedBy", profile.Id);
                        updateCommand.Parameters.AddWithValue("@CurrencyISOCode", pCurrencyModel.CurrencyISOCode);
                        updateCommand.CommandText = UpdateSQL;
                        int RowsAffected = updateCommand.ExecuteNonQuery();
                    }
                    else
                    {
                        TempData["Failed"] = "Error occured: ";
                        Response.Redirect("~/Views/Currency/Index.cshtml");
                    }
                }

                catch (Exception ex)
                {
                }
                //// }
                con.Close();
            }
        }
        #region DropDownTree functions
        private IEnumerable<DropDownTreeItemModel> GetChartOfAccountClassicification()
        {
            List<DropDownTreeItemModel> inlineDefault = new List<DropDownTreeItemModel>();
            inlineDefault = GetChartOfAccountItems("0");
            return inlineDefault;
        }
        private IEnumerable<DropDownTreeItemModel> GetChartOfAccountClassicificationWithHeads()
        {
            List<DropDownTreeItemModel> inlineDefault = new List<DropDownTreeItemModel>();
            inlineDefault = GetChartOfAccountItemsWithHeads("0");
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
        #endregion
        public ActionResult Edit(string id)
        {
            //var accounts = db.ListCurrencies.Where(x => x.CurrencyISOCode == id).Take(1).FirstOrDefault();
            //CurrencyModel currencyModel = new CurrencyModel();
            //currencyModel.CurrencyISOCode = accounts.CurrencyISOCode;
            //currencyModel.CurrencyName = accounts.CurrencyName;
            //currencyModel.AddedBy = accounts.AddedBy;
            //currencyModel.CurrencyRate = Common.toDecimal(accounts.CurrencyRate);
            //currencyModel.CurrencyStatus = Common.toBool(accounts.CurrencyStatus);
            //currencyModel.CurrenyType = Common.toString(accounts.CurrenyType);
            //currencyModel.CurrencySymbol = Common.toString(accounts.CurrencySymbol);
            //currencyModel.LastUpdatedDate = Common.toDateTime(accounts.LastUpdatedBy);
            //currencyModel.isProfit = Common.toBool(accounts.isProfit);
            //currencyModel.MinRateLimit = Common.toDecimal(accounts.MinRateLimit);
            //currencyModel.MaxRateLimit = Common.toDecimal(accounts.MaxRateLimit);
            //currencyModel.CountryISOCode = accounts.CountryISOCode;
            //currencyModel.RevalueYN = Common.toBool(accounts.RevalueYN);
            //currencyModel.CreateHoldAccount = Common.toBool(accounts.AutoCreateUnPaidAccount);
            //currencyModel.CreateAdminFeeAccount = Common.toBool(accounts.AutoCreatUnEarnedAdminAccount);
            //currencyModel.CreateAgentFeeAccount = Common.toBool(accounts.AutoCreateUnEarnedAgentAccount);

            CurrencyModel currencyModel = new CurrencyModel();
            string sql = "select * from ListCurrencies where CurrencyISOCode='" + id + "'";
            DataTable dt = DBUtils.GetDataTable(sql);
            if (dt != null && dt.Rows.Count > 0)
            {
                currencyModel.CurrencyISOCode = Common.toString(dt.Rows[0]["CurrencyISOCode"]);
                currencyModel.CurrencyName = Common.toString(dt.Rows[0]["CurrencyName"]);
                currencyModel.AddedBy = Common.toString(dt.Rows[0]["AddedBy"]);
                currencyModel.CurrencyRate = Common.toDecimal(dt.Rows[0]["CurrencyRate"]);
                currencyModel.CurrencyStatus = Common.toBool(dt.Rows[0]["CurrencyStatus"]);
                currencyModel.CurrenyType = Common.toString(dt.Rows[0]["CurrenyType"]);
                currencyModel.CurrencySymbol = Common.toString(dt.Rows[0]["CurrencySymbol"]);
                currencyModel.LastUpdatedDate = Common.toDateTime(dt.Rows[0]["LastUpdatedBy"]);
                currencyModel.isProfit = Common.toBool(dt.Rows[0]["isProfit"]);
                currencyModel.MinRateLimit = Common.toDecimal(dt.Rows[0]["MinRateLimit"]);
                currencyModel.MaxRateLimit = Common.toDecimal(dt.Rows[0]["MaxRateLimit"]);
                currencyModel.CountryISOCode = Common.toString(dt.Rows[0]["CountryISOCode"]);
                currencyModel.RevalueYN = Common.toBool(dt.Rows[0]["RevalueYN"]);
                currencyModel.CreateHoldAccount = Common.toBool(dt.Rows[0]["AutoCreateUnPaidAccount"]);
                currencyModel.CreateAdminFeeAccount = Common.toBool(dt.Rows[0]["AutoCreatUnEarnedAdminAccount"]);
                currencyModel.CreateAgentFeeAccount = Common.toBool(dt.Rows[0]["AutoCreateUnEarnedAgentAccount"]);
            }
            //currencyModel.UnEarnedAdminAccountCode = accounts.UnEarnedAdminAccountCode;
            //currencyModel.UnEarnedAgentAccountCode = accounts.UnEarnedAgentAccountCode;
            //currencyModel.UnEarnedAdminAccountCode = Common.toString(accounts.UnEarnedAdminAccountCode).Replace("\r\n", "");
            //currencyModel.UnEarnedAdminAccountName = Common.getGLAccountNameCodeByCode(Common.toString(currencyModel.UnEarnedAdminAccountCode)).Replace("\r\n", "");
            //currencyModel.UnEarnedAgentAccountCode = Common.toString(accounts.UnEarnedAgentAccountCode).Replace("\r\n", "");
            //currencyModel.UnEarnedAgentAccountName = Common.getGLAccountNameCodeByCode(Common.toString(currencyModel.UnEarnedAgentAccountCode)).Replace("\r\n", "");
            //currencyModel.HoldAccountCode = Common.toString(accounts.HoldAccountCode).Replace("\r\n", "");
            //currencyModel.HoldAccountName = Common.getGLAccountNameCodeByCode(Common.toString(currencyModel.HoldAccountCode)).Replace("\r\n", "");
            //ViewBag.inlineDefault = GetChartOfAccountClassicificationWithHeads();
            return View("~/Views/Currency/Edit.cshtml", currencyModel);
        }

        [HttpPost]
        public ActionResult Edit(CurrencyModel pCurrencyModel)
        {
            ViewBag.inlineDefault = GetChartOfAccountClassicificationWithHeads();
            if (ModelState.IsValid)
            {
                string UpdateSQL = "Update ListCurrencies Set CurrencyName= @CurrencyName,";
                UpdateSQL += "CurrenyType=@CurrenyType, AutoCreateUnPaidAccount=@AutoCreateUnPaidAccount,AutoCreatUnEarnedAdminAccount=@AutoCreatUnEarnedAdminAccount,AutoCreateUnEarnedAgentAccount=@AutoCreateUnEarnedAgentAccount,CurrencySymbol=@CurrencySymbol,CurrencyStatus=@CurrencyStatus,RevalueYN=@RevalueYN,isProfit=@isProfit,MinRateLimit=@MinRateLimit,MaxRateLimit=@MaxRateLimit,LastUpdatedDate=@LastUpdatedDate,LastUpdatedBy=@LastUpdatedBy Where CurrencyISOCode =@CurrencyISOCode";
                SqlConnection con = Common.getConnection();
                ////using (SqlConnection con = Common.getConnection())
                ////{
                SqlCommand updateCommand = con.CreateCommand();
                try
                {
                    updateCommand.Parameters.AddWithValue("@CurrencyName", pCurrencyModel.CurrencyName);
                    updateCommand.Parameters.AddWithValue("@CurrencySymbol", pCurrencyModel.CurrencySymbol);
                    updateCommand.Parameters.AddWithValue("@CountryISOCode", pCurrencyModel.CountryISOCode);
                    updateCommand.Parameters.AddWithValue("@CurrencyStatus", pCurrencyModel.CurrencyStatus);
                    updateCommand.Parameters.AddWithValue("@RevalueYN", pCurrencyModel.RevalueYN);
                    updateCommand.Parameters.AddWithValue("@isProfit", pCurrencyModel.isProfit);
                    updateCommand.Parameters.AddWithValue("@MinRateLimit", pCurrencyModel.MinRateLimit);
                    updateCommand.Parameters.AddWithValue("@MaxRateLimit", pCurrencyModel.MaxRateLimit);
                    updateCommand.Parameters.AddWithValue("@CurrenyType", pCurrencyModel.CurrenyType);
                    updateCommand.Parameters.AddWithValue("@LastUpdatedDate", DateTime.Now);
                    updateCommand.Parameters.AddWithValue("@LastUpdatedBy", profile.Id);
                    updateCommand.Parameters.AddWithValue("@CurrencyISOCode", pCurrencyModel.CurrencyISOCode);
                    updateCommand.Parameters.AddWithValue("@AutoCreateUnPaidAccount", pCurrencyModel.CreateHoldAccount);
                    updateCommand.Parameters.AddWithValue("@AutoCreatUnEarnedAdminAccount", pCurrencyModel.CreateAdminFeeAccount);
                    updateCommand.Parameters.AddWithValue("@AutoCreateUnEarnedAgentAccount", pCurrencyModel.CreateAgentFeeAccount);
                    updateCommand.CommandText = UpdateSQL;
                    int RowsAffected = updateCommand.ExecuteNonQuery();
                    if (RowsAffected == 1)
                    {
                        TempData["Success"] = "Currency updated successfully";
                        return View("~/Views/Currency/Edit.cshtml", pCurrencyModel);
                    }
                    else
                    {
                        TempData["Failed"] = "Error occured: ";
                        return View("~/Views/Currency/Edit.cshtml", pCurrencyModel);
                    }

                }
                catch (Exception ex)
                {
                    TempData["Failed"] = "Error occured: " + ex.Message.ToString();
                    return View("~/Views/Currency/Edit.cshtml", pCurrencyModel);
                }
                ////  }
                con.Close();
            }
            else
            {
                TempData["Failed"] = "Data missing. Please check all fields and try again";
                return View("~/Views/Currency/Edit.cshtml", pCurrencyModel);
            }
        }

        public ActionResult Delete(string id)
        {
            //var accounts = db.ListCurrencies.Find(id);
            //CurrencyModel currencyModel = new CurrencyModel();
            //currencyModel.CurrencyISOCode = accounts.CurrencyISOCode;
            //currencyModel.CurrencyName = accounts.CurrencyName;
            //currencyModel.AddedBy = accounts.AddedBy;
            //currencyModel.CurrencyRate = Common.toDecimal(accounts.CurrencyRate);
            //currencyModel.CurrencyStatus = Common.toBool(accounts.CurrencyStatus);
            //currencyModel.CurrencySymbol = Common.toString(accounts.CurrencySymbol);
            //currencyModel.DateAdded = Common.toDateTime(accounts.DateAdded);
            //currencyModel.HoldAccountCode = Common.toString(accounts.HoldAccountCode);
            //currencyModel.UnEarnedAdminAccountCode = accounts.UnEarnedAdminAccountCode;
            //currencyModel.UnEarnedAgentAccountCode = accounts.UnEarnedAgentAccountCode;
            CurrencyModel currencyModel = new CurrencyModel();
            string sql = "select * from ListCurrencies where CurrencyISOCode='" + id + "'";
            DataTable dt = DBUtils.GetDataTable(sql);
            if (dt != null && dt.Rows.Count > 0)
            {
                currencyModel.CurrencyISOCode = Common.toString(dt.Rows[0]["CurrencyISOCode"]);
                currencyModel.CurrencyName = Common.toString(dt.Rows[0]["CurrencyName"]);
                currencyModel.AddedBy = Common.toString(dt.Rows[0]["AddedBy"]);
                currencyModel.CurrencyRate = Common.toDecimal(dt.Rows[0]["CurrencyRate"]);
                currencyModel.CurrencyStatus = Common.toBool(dt.Rows[0]["CurrencyStatus"]);
                currencyModel.CurrencySymbol = Common.toString(dt.Rows[0]["CurrencySymbol"]);
                currencyModel.DateAdded = Common.toDateTime(dt.Rows[0]["DateAdded"]);
                currencyModel.HoldAccountCode = Common.toString(dt.Rows[0]["HoldAccountCode"]);
                currencyModel.UnEarnedAdminAccountCode = Common.toString(dt.Rows[0]["UnEarnedAdminAccountCode"]);
                currencyModel.UnEarnedAgentAccountCode = Common.toString(dt.Rows[0]["UnEarnedAgentAccountCode"]);
            }

            return View("~/Views/Currency/Delete.cshtml", currencyModel);
        }

        [HttpPost]
        public ActionResult Delete(CurrencyModel pCurrencyModel)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    //string sql = "select top 1 AccountCode from GLChartOfAccounts where AccountID = '" + chartOfAccounts.AccountID + "' order by AccountID desc";
                    //string IsChild = DBUtils.executeSqlGetSingle(sql);

                    //if (string.IsNullOrEmpty(IsChild))
                    //{
                    string sqlDelete = "Delete From ListCurrencies Where CurrencyISOCode = '" + pCurrencyModel.CurrencyISOCode + "'";
                    bool Isfound = DBUtils.ExecuteSQL(sqlDelete);
                    if (Isfound)
                    {
                        TempData["Success"] = "Successfully deleted.";
                        return View("~/Views/Currency/delete.cshtml", pCurrencyModel);
                    }
                    else
                    {
                        TempData["Failed"] = "Error occured: ";
                        return View("~/Views/Currency/delete.cshtml", pCurrencyModel);
                    }
                    //}
                    //else
                    //{
                    //    TempData["Failed"] = "Error occured: ";
                    //    return View("~/Views/ChartOfAccounts/hdelete.cshtml", chartOfAccounts);
                    //}
                }
                catch (Exception ex)
                {
                    TempData["Failed"] = "Error occured: " + ex.Message.ToString();
                    return View("~/Views/Currency/delete.cshtml", pCurrencyModel);
                }
            }
            else
            {
                TempData["Failed"] = "Invalid account id.";
                return View("~/Views/Currency/delete.cshtml", pCurrencyModel);
            }
        }
        #endregion

        #region Rate Management
        public ActionResult RIndex()
        {
            return View("~/Views/Currency/RIndex.cshtml");
        }


        public ActionResult RCheckList_Read([DataSourceRequest] DataSourceRequest request)
        {
            string countQuery = @"SELECT COUNT(5) FROM ListCurrencies /**where**/";
            string selectQuery = @"SELECT  *FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, ";
            selectQuery += @" CurrencyISOCode,CurrencyName ,CurrencySymbol, (select top 1 CountryName from ListCountries where ListCountries.CountryISONumericCode=ListCurrencies.CountryISOCode ) as  CountryISOCode, CurrencyStatus, MinRateLimit, MaxRateLimit, CurrenyType, FORMAT ( LastUpdatedDate, 'd', 'en-gb' ) as LastUpdatedDate,(select top 1 ExchangeRate from ExchangeRates where ExchangeRates.CurrencyISOCode=ListCurrencies.CurrencyISOCode and AddedDate = '" + Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00") + "') as  CurrencyRate,(select top 1 RevalueExchangeRate from ExchangeRates where ExchangeRates.CurrencyISOCode=ListCurrencies.CurrencyISOCode and AddedDate = '" + Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00") + "') as  RevalueExchangeRate FROM ListCurrencies";
            selectQuery += @" /**where**/ ) AS RowConstrainedResult WHERE   RowNum >= (@PageIndex * @PageSize + 1 ) AND RowNum <= (@PageIndex + 1) * @PageSize ORDER BY RowNum";
            //     const string selectQuery = @"SELECT
            //CheckListId, CheckTypeId, (select CheckType  from CheckTypes where CheckTypes.CheckTypeId = CheckList.CheckTypeId ) as CheckTypeName, CheckText, case when Status = 'true' then 'Active' else 'Inactive' end as StatusName, Status FROM CheckList";
            SqlBuilder builder = new SqlBuilder();
            var count = builder.AddTemplate(countQuery);

            int CurrentPage = 0; int CurrentPageSize = 25;
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
            //else
            //{
            //    builder.Where("CurrencyStatus=1");
            //}


            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("CurrencyISOCode");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);

            var rows = _dbcontext.Query<CurrencyModel>(selector.RawSql, selector.Parameters);

            //  rows.Each(item => Console.WriteLine($"({ item.AccountCode}): { item.AccountName}
            //{                GetParentsString(items, item)}"));            
            //}
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult REditingInline_Update([DataSourceRequest] DataSourceRequest request, CurrencyModel product)
        {
            if (product != null)
            {
                REditingCustom_Update(product);
            }

            return Json(new[] { product }.ToDataSourceResult(request, ModelState));
        }
        public void REditingCustom_Update(CurrencyModel pCurrencyModel)
        {
            if (pCurrencyModel.CurrencyISOCode != "" && pCurrencyModel.MinRateLimit > 0.0001m && pCurrencyModel.MinRateLimit > 0.0001m)
            {
                if (pCurrencyModel.CurrencyRate >= pCurrencyModel.MinRateLimit && pCurrencyModel.CurrencyRate <= pCurrencyModel.MaxRateLimit)
                {
                    string ExchangeRateID = DBUtils.executeSqlGetSingle("Select ExchangeRateID from ExchangeRates Where CurrencyISOCode='" + pCurrencyModel.CurrencyISOCode + "' and AddedDate='" + Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00") + "'");
                    if (Common.toString(ExchangeRateID).Trim() != "")
                    {
                        string UpdateSQL = "Update ExchangeRates Set ";
                        UpdateSQL += " ExchangeRate=@ExchangeRate,RevalueExchangeRate=@RevalueExchangeRate, UpdatedDate=@UpdatedDate, UpdatedBy=@UpdatedBy Where ExchangeRateID = @ExchangeRateID";
                        SqlConnection con = Common.getConnection();
                        ////using (SqlConnection con = Common.getConnection())
                        ////{
                        SqlCommand updateCommand = con.CreateCommand();
                        try
                        {
                            updateCommand.Parameters.AddWithValue("@ExchangeRate", pCurrencyModel.CurrencyRate);
                            updateCommand.Parameters.AddWithValue("@RevalueExchangeRate", pCurrencyModel.CurrencyRate);
                            updateCommand.Parameters.AddWithValue("@UpdatedDate", Common.toDateTime(DateTime.Now));
                            updateCommand.Parameters.AddWithValue("@UpdatedBy", profile.Id);
                            updateCommand.Parameters.AddWithValue("@ExchangeRateID", ExchangeRateID);
                            updateCommand.CommandText = UpdateSQL;
                            int RowsAffected = updateCommand.ExecuteNonQuery();
                            updateCommand.Parameters.Clear();
                            if (RowsAffected > 0)
                            {
                                string UpdateListCurreny = "Update ListCurrencies set LastUpdatedDate='" + DateTime.Now + "' where CurrencyISOCode='" + pCurrencyModel.CurrencyISOCode + "'";
                                updateCommand.CommandText = UpdateListCurreny;
                                updateCommand.ExecuteNonQuery();
                                updateCommand.Parameters.Clear();
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                        ////}
                        con.Close();
                    }
                    else
                    {
                        string UpdateSQL = "Insert into ExchangeRates (CurrencyISOCode,ExchangeRate,AddedDate,AddedBy, UpdatedDate, UpdatedBy, RevalueExchangeRate) values (@Currency, @ExchangeRate, @AddedDate, @AddedBy, @UpdatedDate, @UpdatedBy, @RevalueExchangeRate) ";
                        SqlConnection con = Common.getConnection();
                        ////using (SqlConnection con = Common.getConnection())
                        ////{
                        SqlCommand InsertCommand = con.CreateCommand();
                        try
                        {
                            InsertCommand.Parameters.AddWithValue("@ExchangeRate", pCurrencyModel.CurrencyRate);
                            InsertCommand.Parameters.AddWithValue("@Currency", pCurrencyModel.CurrencyISOCode);
                            InsertCommand.Parameters.AddWithValue("@AddedDate", Common.toDateTime(SysPrefs.PostingDate));
                            InsertCommand.Parameters.AddWithValue("@AddedBy", profile.Id);
                            InsertCommand.Parameters.AddWithValue("@UpdatedDate", DateTime.Now);
                            InsertCommand.Parameters.AddWithValue("@UpdatedBy", profile.Id);
                            InsertCommand.Parameters.AddWithValue("@RevalueExchangeRate", pCurrencyModel.CurrencyRate);
                            InsertCommand.CommandText = UpdateSQL;
                            int RowsAffected = InsertCommand.ExecuteNonQuery();
                            if (RowsAffected > 0)
                            {
                                string UpdateListCurreny = "Update ListCurrencies set LastUpdatedDate='" + DateTime.Now + "' where CurrencyISOCode='" + pCurrencyModel.CurrencyISOCode + "'";
                                InsertCommand.CommandText = UpdateListCurreny;
                                InsertCommand.ExecuteNonQuery();
                                InsertCommand.Parameters.Clear();
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                        ////}
                        con.Close();
                    }
                }
                //if (pCurrencyModel.RevalueExchangeRate >= pCurrencyModel.MinRateLimit && pCurrencyModel.RevalueExchangeRate <= pCurrencyModel.MaxRateLimit)
                //{
                //    string ExchangeRateID = DBUtils.executeSqlGetSingle("Select ExchangeRateID from ExchangeRates Where CurrencyISOCode='" + pCurrencyModel.CurrencyISOCode + "' and AddedDate='" + Common.toDateTime(SysPrefs.PostingDate).ToString("yyyy-MM-dd 00:00:00") + "'");
                //    if (Common.toString(ExchangeRateID).Trim() != "")
                //    {
                //        string UpdateSQL = "Update ExchangeRates Set ";
                //        UpdateSQL += " RevalueExchangeRate=@ExchangeRate, UpdatedDate=@UpdatedDate, UpdatedBy=@UpdatedBy Where ExchangeRateID = @ExchangeRateID";
                //        SqlConnection con = Common.getConnection();
                //        //using (SqlConnection con = Common.getConnection())
                //        //{
                //        SqlCommand updateCommand = con.CreateCommand();
                //        try
                //        {
                //            updateCommand.Parameters.AddWithValue("@ExchangeRate", pCurrencyModel.RevalueExchangeRate);
                //            updateCommand.Parameters.AddWithValue("@UpdatedDate", Common.toDateTime(DateTime.Now));
                //            updateCommand.Parameters.AddWithValue("@UpdatedBy", profile.Id);
                //            updateCommand.Parameters.AddWithValue("@ExchangeRateID", ExchangeRateID);
                //            updateCommand.CommandText = UpdateSQL;
                //            int RowsAffected = updateCommand.ExecuteNonQuery();
                //            updateCommand.Parameters.Clear();
                //            if (RowsAffected > 0)
                //            {
                //                string UpdateListCurreny = "Update ListCurrencies set LastUpdatedDate='" + DateTime.Now + "' where CurrencyISOCode='" + pCurrencyModel.CurrencyISOCode + "'";
                //                updateCommand.CommandText = UpdateListCurreny;
                //                updateCommand.ExecuteNonQuery();
                //                updateCommand.Parameters.Clear();
                //            }
                //        }
                //        catch (Exception ex)
                //        {
                //        }
                //        //}
                //        con.Close();
                //    }
                //    else
                //    {
                //        string UpdateSQL = "Insert into ExchangeRates (CurrencyISOCode,ExchangeRate,AddedDate,AddedBy, UpdatedDate, UpdatedBy, RevalueExchangeRate) values (@Currency, @ExchangeRate, @AddedDate, @AddedBy, @UpdatedDate, @UpdatedBy, @RevalueExchangeRate) ";
                //        SqlConnection con = Common.getConnection();
                //        //using (SqlConnection con = Common.getConnection())
                //        //{
                //        SqlCommand InsertCommand = con.CreateCommand();
                //        try
                //        {
                //            InsertCommand.Parameters.AddWithValue("@ExchangeRate", pCurrencyModel.CurrencyRate);
                //            InsertCommand.Parameters.AddWithValue("@Currency", pCurrencyModel.CurrencyISOCode);
                //            InsertCommand.Parameters.AddWithValue("@AddedDate", Common.toDateTime(SysPrefs.PostingDate));
                //            InsertCommand.Parameters.AddWithValue("@AddedBy", profile.Id);
                //            InsertCommand.Parameters.AddWithValue("@UpdatedDate", DateTime.Now);
                //            InsertCommand.Parameters.AddWithValue("@UpdatedBy", profile.Id);
                //            InsertCommand.Parameters.AddWithValue("@RevalueExchangeRate", pCurrencyModel.RevalueExchangeRate);
                //            InsertCommand.CommandText = UpdateSQL;
                //            int RowsAffected = InsertCommand.ExecuteNonQuery();
                //            if (RowsAffected > 0)
                //            {
                //                string UpdateListCurreny = "Update ListCurrencies set LastUpdatedDate='" + DateTime.Now + "' where CurrencyISOCode='" + pCurrencyModel.CurrencyISOCode + "'";
                //                InsertCommand.CommandText = UpdateListCurreny;
                //                InsertCommand.ExecuteNonQuery();
                //                InsertCommand.Parameters.Clear();
                //            }
                //        }
                //        catch (Exception ex)
                //        {
                //        }
                //        //}
                //        con.Close();
                //    }
                //}
                //else
                //{
                //    TempData["Failed"] = "Error occured: ";
                //    Response.Redirect("~/Views/Currency/RIndex.cshtml");
                //}
            }
        }
        #endregion
    }
}