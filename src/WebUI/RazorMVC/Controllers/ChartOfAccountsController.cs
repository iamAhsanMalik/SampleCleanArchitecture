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
    public class ChartOfAccountsController : Controller
    {
        //// private iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
        ApplicationUser Profile = ApplicationUser.GetUserProfile();
        private AppDbContext _dbcontext = new AppDbContext();
        //// private AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());
        #region Chart Of Accounts Classifications

        public ActionResult cIndex()
        {
            return View("~/Views/ChartOfAccounts/cIndex.cshtml");
        }
        public JsonResult TreeList_Read([DataSourceRequest] DataSourceRequest request)
        {
            // var result = db.GLChartOfAccounts.Where(x => x.isHead == false).OrderBy(x => x.AccountCode).ToTreeDataSourceResult(request,
            //    e => e.AccountCode,
            //    e => e.ParentAccountCode,
            //    e => e
            //);
            var result = _dbcontext.Query<ChartOfAccountsModel>("select AccountID,AccountName,AccountCode,ParentAccountCode from GLChartOfAccounts where isHead= 0 order by AccountCode").ToTreeDataSourceResult(request,
                e => e.AccountCode,
                e => e.ParentAccountCode,
                e => e
            );
            return Json(result, JsonRequestBehavior.AllowGet);
        }
        private IEnumerable<ChartOfAccountsModel> TreeList()
        {
            var data = _dbcontext.Query<ChartOfAccountsModel>("select * from GLChartOfAccounts where isHead= 0");
            return data.AsEnumerable();
        }
        public JsonResult Accounts_Read(string id)
        {
            #region 
            //List<DropDownTreeItemModel> inlineDefault = new List<DropDownTreeItemModel>
            //    {
            //        new DropDownTreeItemModel
            //        {
            //            Text = "Furniture",
            //            Id="1",
            //            Items = new List<DropDownTreeItemModel>
            //            {
            //                new DropDownTreeItemModel()
            //                {
            //                    Text = "Tables & Chairs"
            //                },
            //                new DropDownTreeItemModel
            //                {
            //                     Text = "Sofas"
            //                }
            //            }
            //        }
            //};

            //IEnumerable<DropDownTreeItemModel> result;

            //if (string.IsNullOrEmpty(id))
            //{
            //    result = TreeViewRepository.GetProjectData().Select(o => o.Clone());
            //}
            //else
            //{
            //    result = TreeViewRepository.GetChildren(id);
            //}

            //return Json(result, JsonRequestBehavior.AllowGet);
            #endregion

            List<DropDownTreeItemModel> result = new List<DropDownTreeItemModel>();
            DropDownTreeItemModel item = new DropDownTreeItemModel();
            item.Id = "100";
            item.Text = "100";
            result.Add(item);

            #region 
            //result = TreeList;
            //using (db)
            //{

            //    var accounts = db.GLChartOfAccounts

            //                                .Where(account => AccountId.HasValue ? account.AccountID == AccountId : account.AccountCode == null);

            //    var result = accounts.Select(account => new
            //    {
            //        AccountId = account.AccountID,
            //        Name = account.AccountName,
            //        HasChildren = account.ParentAccountCode.Any()
            //    })
            //                            .ToList();
            //    return Json(result, JsonRequestBehavior.AllowGet);
            //}
            //var employees = from e in db.GLChartOfAccounts
            //                where (e.ParentAccountCode == "0")
            //                select new
            //                {
            //                    id = e.AccountCode,
            //                    Name = e.AccountName,
            //                    hasChildren = e.ParentAccountCode.Any()
            //                };
            //return Json(employees, JsonRequestBehavior.AllowGet);

            // IEnumerable<ChartOfAccountsModel> result;

            //result = from e in db.GLChartOfAccounts
            //         where (e.ParentAccountCode == "0")
            //         select new ChartOfAccountsModel
            //         {
            //             AccountCode = e.AccountCode,
            //             AccountName = e.AccountName,
            //             ParentAccountCode = e.ParentAccountCode
            //         };
            #endregion
            return Json(result, JsonRequestBehavior.AllowGet);
        }
        public ActionResult cCreate(string Id)
        {
            ChartOfAccountsModel pModel = new ChartOfAccountsModel();
            int myId = Common.toInt(Id);
            //  var getlist = db.GLChartOfAccounts.Where(x => x.AccountID == myId).FirstOrDefault();
            string sql = "select AccountCode,AccountName from GLChartOfAccounts where AccountID=" + myId + "";
            DataTable getlist = DBUtils.GetDataTable(sql);
            if (getlist != null)
            {
                if (getlist.Rows.Count > 0)
                {
                    DataRow drChartofAccounts = getlist.Rows[0];
                    string getAccountCount = Common.toString(DBUtils.executeSqlGetSingle("select Count(*) from GLChartOfAccounts where ParentAccountCode=" + drChartofAccounts["AccountCode"].ToString() + ""));
                    if (getAccountCount == "")
                    {
                        getAccountCount = "0";
                    }
                    if (Common.toInt(getAccountCount) <= 9)
                    {
                        pModel.ParentAccountName = Common.toString(drChartofAccounts["AccountName"]); //getlist.AccountName;
                        pModel.ParentAccountCode = Common.toString(drChartofAccounts["AccountCode"]); // Common.toString(getlist.AccountCode);

                    }
                    else
                    {
                        pModel.isVisible = true;
                        TempData["Failed"] = "Already 10 subAccount exist for this Account";
                    }
                }
            }
            pModel.AccountStatus = true;
            return View("~/Views/ChartOfAccounts/cCreate.cshtml", pModel);
        }
        [HttpPost]
        public ActionResult cCreate(ChartOfAccountsModel model)
        {
            if (ModelState.IsValid)
            {
                #region old code
                //string code1 = Common.toString(model.ParentAccountCode) + "1";
                //string SqlGetlastChildAccountCode = "select top 1 AccountCode  from GLChartOfAccounts where ParentAccountCode = '" + model.ParentAccountCode + "' order by AccountID desc";
                //string LastChildAccountCode = DBUtils.executeSqlGetSingle(SqlGetlastChildAccountCode);
                //if (!string.IsNullOrEmpty(LastChildAccountCode)) {
                //    code1 = LastChildAccountCode;
                //}
                //else
                //{
                //    code1 = Common.toString(model.ParentAccountCode) + "0";
                //}
                #endregion                                

                string tempAccountCode = "";
                string getAccountCount = Common.toString(DBUtils.executeSqlGetSingle("select Count(*) from GLChartOfAccounts where ParentAccountCode='" + model.ParentAccountCode + "'"));
                if (getAccountCount == "")
                {
                    getAccountCount = "0";
                }
                if (Common.toInt(getAccountCount) <= 9)
                {
                    for (int i = 1; i <= 9; i++)
                    {
                        if (model.ParentAccountCode.Length == 4)
                        {
                            tempAccountCode = model.ParentAccountCode + Common.toString(i).PadLeft(4, '0');
                        }
                        else
                        {
                            tempAccountCode = model.ParentAccountCode + Common.toString(i);
                        }
                        string sql = "select AccountCode  from GLChartOfAccounts where  AccountCode = '" + tempAccountCode + "' and ParentAccountCode = '" + model.ParentAccountCode + "' order by AccountCode asc";
                        string pCode = DBUtils.executeSqlGetSingle(sql);
                        if (string.IsNullOrEmpty(pCode))
                        {
                            // code1 = Common.toString(Common.toInt(code1) + 1);
                            model.AccountCode = tempAccountCode;
                            break;
                        }
                    }
                    #region old code
                    // SqlGetlastChildAccountCode = "select top 1 AccountCode  from GLChartOfAccounts where ParentAccountCode = '" + model.ParentAccountCode + "' order by AccountID desc";
                    // LastChildAccountCode = DBUtils.executeSqlGetSingle(SqlGetlastChildAccountCode);
                    //if (!string.IsNullOrEmpty(LastChildAccountCode))
                    //{
                    //    LastChildAccountCode = (1 + Common.toInt(LastChildAccountCode)).ToString();
                    //}
                    //else
                    //{
                    //    LastChildAccountCode = Common.toString(model.ParentAccountCode) + "1";
                    //}
                    //   model.AccountCode = LastChildAccountCode;
                    #endregion
                    //model.AccountCode = code1;
                    string SqlInsert = "Insert Into GLChartOfAccounts (AccountCode, AccountName, ParentAccountCode, AddedBy ,AddedDate, AccountStatus,isHead)";
                    SqlInsert += "Values(@AccountCode, @AccountName, @ParentAccountCode, @AddedBy, @AddedDate, @AccountStatus,@isHead)";
                    bool myStatus = model.AccountStatus;
                    SqlConnection con = Common.getConnection();
                    ////using (SqlConnection con = Common.getConnection())
                    ////{
                    SqlCommand InsertCommand = con.CreateCommand();
                    try
                    {
                        InsertCommand.Parameters.AddWithValue("@AccountCode", model.AccountCode);
                        InsertCommand.Parameters.AddWithValue("@AccountName", model.AccountName);
                        InsertCommand.Parameters.AddWithValue("@ParentAccountCode", model.ParentAccountCode);
                        InsertCommand.Parameters.AddWithValue("@AddedBy", Profile.Id);
                        InsertCommand.Parameters.AddWithValue("@AddedDate", DateTime.Now);
                        InsertCommand.Parameters.AddWithValue("@AccountStatus", myStatus);
                        InsertCommand.Parameters.AddWithValue("@isHead", 0);
                        InsertCommand.CommandText = SqlInsert;
                        int RowsAffected = InsertCommand.ExecuteNonQuery();

                        if (RowsAffected == 1)
                        {
                            TempData["Success"] = "Successfully Created Sub Account";
                            return View("~/Views/ChartOfAccounts/cCreate.cshtml", model);
                        }
                        else
                        {
                            TempData["Failed"] = " Failed to add. Please try again";
                            return View("~/Views/ChartOfAccounts/cCreate.cshtml", model);
                        }
                    }
                    catch (Exception ex)
                    {
                        TempData["Failed"] = " Failed to add. Please try again" + ex.Message.ToString();
                        return View("~/Views/ChartOfAccounts/cCreate.cshtml", model);
                    }
                    finally
                    {
                        con.Close();
                    }
                    //// }

                }
                else
                {
                    TempData["Failed"] = "Already 10 subAccount exist for this Account";
                    return View("~/Views/ChartOfAccounts/cCreate.cshtml", model);
                }
            }
            else
            {
                TempData["Failed"] = " Failed to add. Please try again";
                return View("~/Views/ChartOfAccounts/cCreate.cshtml", model);
            }
        }
        public ActionResult cEdit(string Id)
        {
            ChartOfAccountsModel pModel = new ChartOfAccountsModel();
            int myId = Common.toInt(Id);
            //  var getlist = db.GLChartOfAccounts.Where(x => x.AccountID == myId).FirstOrDefault();
            //  if (getlist != null)
            string sql = "select AccountID,AccountCode,AccountName,ParentAccountCode,AccountStatus from GLChartOfAccounts where AccountID=" + myId + "";
            DataTable dt = DBUtils.GetDataTable(sql);
            if (dt != null && dt.Rows.Count > 0)
            {
                pModel.AccountName = Common.toString(dt.Rows[0]["AccountName"]);  //getlist.AccountName;
                pModel.AccountID = Common.toInt(dt.Rows[0]["AccountID"]); //getlist.AccountID;
                pModel.AccountCode = Common.toString(dt.Rows[0]["AccountCode"]); //Common.toString(getlist.AccountCode);
                pModel.ParentAccountCode = Common.toString(dt.Rows[0]["ParentAccountCode"]); //Common.toString(getlist.ParentAccountCode);
                pModel.AccountStatus = Common.toBool(dt.Rows[0]["AccountStatus"]); //Common.toBool(getlist.AccountStatus);
            }
            return View("~/Views/ChartOfAccounts/cEdit.cshtml", pModel);
        }
        [HttpPost]
        public ActionResult cEdit(ChartOfAccountsModel model)
        {
            if (ModelState.IsValid)
            {
                //string UpdateSQL = "Update GLChartOfAccounts Set AccountCode=@AccountCode,ParentAccountCode=@ParentAccountCode, AccountName= @AccountName,TypeID=@TypeID,";
                //UpdateSQL += "AccountStatus=@AccountStatus,[Prefix] = @Prefix, [CurrencyISOCode] = @CurrencyISOCode,[CountryISONumericCode] = @CountryISONumericCode, ";
                //UpdateSQL += " [RevalueYN] = @RevalueYN, [Revalue] = @Revalue, [DebiteLimit] = @DebiteLimit, [CreditLimit] = @CreditLimit, [TaxPercentage] = @TaxPercentage";
                //UpdateSQL += ",[VATWHT] = @VATWHT,[Charges] = @Charges  Where AccountID =@AccountID";

                string UpdateSQL = "Update GLChartOfAccounts Set AccountCode=@AccountCode,ParentAccountCode=@ParentAccountCode, AccountName= @AccountName,";
                UpdateSQL += "AccountStatus=@AccountStatus  Where AccountID =@AccountID";

                ////using (SqlConnection con = Common.getConnection())
                ////{
                SqlConnection con = Common.getConnection();
                SqlCommand updateCommand = con.CreateCommand();
                try
                {
                    updateCommand.Parameters.AddWithValue("@AccountCode", model.AccountCode);
                    updateCommand.Parameters.AddWithValue("@AccountName", model.AccountName);
                    updateCommand.Parameters.AddWithValue("@ParentAccountCode", model.ParentAccountCode);
                    updateCommand.Parameters.AddWithValue("@AccountStatus", model.AccountStatus);
                    updateCommand.Parameters.AddWithValue("@AccountID", model.AccountID);

                    updateCommand.CommandText = UpdateSQL;
                    int RowsAffected = updateCommand.ExecuteNonQuery();

                    if (RowsAffected == 1)
                    {
                        TempData["Success"] = " Successfully Updated Sub Account ";
                        return View("~/Views/ChartOfAccounts/cEdit.cshtml", model);
                    }
                    else
                    {
                        TempData["Failed"] = " Failed to Updated. Please try again";
                        return View("~/Views/ChartOfAccounts/cEdit.cshtml", model);
                    }
                }
                catch
                {
                    TempData["Failed"] = " Failed to Updated. Please try again";
                    return View("~/Views/ChartOfAccounts/cEdit.cshtml", model);
                }
                finally
                {
                    con.Close();
                }
            }
            return View("~/Views/ChartOfAccounts/cEdit.cshtml", model);
        }
        public ActionResult cDelete(string Id)
        {
            ChartOfAccountsModel pModel = new ChartOfAccountsModel();
            int myId = Common.toInt(Id);
            //  var getlist = db.GLChartOfAccounts.Where(x => x.AccountID == myId).FirstOrDefault();
            //  if (getlist != null)
            //{
            //    pModel.AccountName = getlist.AccountName;
            //    pModel.AccountID = getlist.AccountID;
            //    pModel.AccountCode = Common.toString(getlist.AccountCode);
            //    pModel.ParentAccountCode = Common.toString(getlist.ParentAccountCode);
            //    pModel.AccountStatus = Common.toBool(getlist.AccountStatus);
            //}
            string sql = "select AccountID,AccountCode,AccountName,ParentAccountCode,AccountStatus from GLChartOfAccounts where AccountID=" + myId + "";
            DataTable dt = DBUtils.GetDataTable(sql);
            if (dt != null && dt.Rows.Count > 0)
            {
                pModel.AccountName = Common.toString(dt.Rows[0]["AccountName"]);  //getlist.AccountName;
                pModel.AccountID = Common.toInt(dt.Rows[0]["AccountID"]); //getlist.AccountID;
                pModel.AccountCode = Common.toString(dt.Rows[0]["AccountCode"]); //Common.toString(getlist.AccountCode);
                pModel.ParentAccountCode = Common.toString(dt.Rows[0]["ParentAccountCode"]); //Common.toString(getlist.ParentAccountCode);
                pModel.AccountStatus = Common.toBool(dt.Rows[0]["AccountStatus"]); //Common.toBool(getlist.AccountStatus);
            }
            return View("~/Views/ChartOfAccounts/cDelete.cshtml", pModel);
        }
        [HttpPost]
        public ActionResult cDelete(ChartOfAccountsModel model)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    string transactionID = DBUtils.executeSqlGetSingle("select TransactionID from GLTransactions where AccountCode='" + model.AccountCode + "'");
                    if (string.IsNullOrEmpty(Common.toString(transactionID)))
                    {
                        string sql = "select top 1 AccountCode from GLChartOfAccounts where ParentAccountCode = '" + model.AccountCode + "' order by AccountID desc";
                        string IsChild = DBUtils.executeSqlGetSingle(sql);

                        if (string.IsNullOrEmpty(IsChild))
                        {
                            string sqlDelete = "Delete From GLChartOfAccounts Where AccountID = '" + model.AccountID + "'";
                            bool Isfound = DBUtils.ExecuteSQL(sqlDelete);
                            if (Isfound)
                            {
                                TempData["Success"] = " Successfully Deleted Sub Account";
                                return View("~/Views/ChartOfAccounts/cDelete.cshtml", model);
                            }
                            else
                            {
                                TempData["Failed"] = " Failed to Delete. Please try again";
                                return View("~/Views/ChartOfAccounts/cDelete.cshtml", model);
                            }
                        }
                        else
                        {
                            TempData["Failed"] = " Failed to Delete. Please try to delete first Sub Accounts";
                            return View("~/Views/ChartOfAccounts/cDelete.cshtml", model);
                        }
                    }
                    else
                    {
                        TempData["Failed"] = " Failed to Delete. This account have some transactions in the database";
                        return View("~/Views/ChartOfAccounts/cDelete.cshtml", model);
                    }
                }
                catch
                {
                    TempData["Failed"] = " Failed to Delete. Please try again";
                    return View("~/Views/ChartOfAccounts/cDelete.cshtml", model);
                }
            }
            return View("~/Views/ChartOfAccounts/cDelete.cshtml", model);
        }
        #endregion

        #region Chart Of Accounts Heads
        // GET: ChartOfAccounts
        public ActionResult hIndex()
        {
            //ViewBag.inlineDefault = Local_Data_Binding_Get_Default_Inline_Data();

            DropDownTreeItemModel inlineDefault = new DropDownTreeItemModel();
            ViewBag.inlineDefault = GetChartOfAccountClassicification();
            return View("~/Views/ChartOfAccounts/hIndex.cshtml");
        }
        private List<DropDownTreeItemModel> GetChartOfAccountItems(string pAccountCode)
        {
            //  ; WITH UpdateParentAccountBalances AS(select  AccountName, AccountCode, ParentAccountCode, isHead from GLChartOfAccounts where ParentAccountCode = '0' union all
            //  select e.AccountName, e.AccountCode, e.ParentAccountCode, e.isHead from GLChartOfAccounts e
            //     INNER JOIN UpdateParentAccountBalances o ON e.ParentAccountCode = o.AccountCode)
            //select distinct *from UpdateParentAccountBalances where (isHead is null or isHead = 0) order by AccountCode;
            List<DropDownTreeItemModel> pList = new List<DropDownTreeItemModel>();
            if (pAccountCode != "")
            {
                string sqlAccount = "select * from GLChartOfAccounts where (isHead is null or isHead=0) And ParentAccountCode='" + pAccountCode + "'";

                //sqlAccount = ";WITH UpdateParentAccountBalances AS(select  AccountName, AccountCode, ParentAccountCode, isHead from GLChartOfAccounts where ParentAccountCode = '0' union all ";
                //sqlAccount += " select e.AccountName, e.AccountCode, e.ParentAccountCode, e.isHead from GLChartOfAccounts e";
                //sqlAccount += " INNER JOIN UpdateParentAccountBalances o ON e.ParentAccountCode = o.AccountCode)";
                //sqlAccount += " select distinct *from UpdateParentAccountBalances where (isHead is null or isHead = 0) order by AccountCode;";


                DataTable dtAccount = DBUtils.GetDataTable(sqlAccount);
                foreach (DataRow dr in dtAccount.Rows)
                {
                    DropDownTreeItemModel item = new DropDownTreeItemModel();
                    item.Text = dr["AccountName"].ToString();
                    item.Id = dr["AccountCode"].ToString();
                    //List<DropDownTreeItemModel> pItems = GetChartOfAccountItems(pList, dr["AccountCode"].ToString());
                    //for (int i = 0; i < pItems.Count; i++)
                    //{
                    //    item.Items.Add(pItems[i]);
                    //}
                    //item.Items.Add(item);
                    item.Expanded = true;
                    item.Items = GetChartOfAccountItems(dr["AccountCode"].ToString());

                    pList.Add(item);
                }
            }
            return pList;
        }
        private IEnumerable<DropDownTreeItemModel> GetChartOfAccountClassicification()
        {
            List<DropDownTreeItemModel> inlineDefault = new List<DropDownTreeItemModel>();
            inlineDefault = GetChartOfAccountItems("0");
            return inlineDefault;

        }
        private IEnumerable<DropDownTreeItemModel> Local_Data_Binding_Get_Default_Inline_Data()
        {
            string sql = "select * from GLChartOfAccounts where ParentAccountCode='0'";
            DataTable dtsql = DBUtils.GetDataTable(sql);
            List<DropDownTreeItemModel> inlineDefault = new List<DropDownTreeItemModel>();
            foreach (DataRow dr in dtsql.Rows)
            {
                sql = "select AccountName,AccountCode from GLChartOfAccounts where ParentAccountCode='" + dr["AccountCode"].ToString() + "'";
                DataTable dtsql1 = DBUtils.GetDataTable(sql);
                foreach (DataRow drsql1 in dtsql1.Rows)
                {
                    new DropDownTreeItemModel
                    {
                        Text = dr["AccountName"].ToString(), //"Furniture",
                        Id = dr["AccountCode"].ToString(),  //"100",
                        Items = new List<DropDownTreeItemModel>
                        {
                            new DropDownTreeItemModel()
                            {
                                Text =drsql1["AccountName"].ToString(),      // "Tables & Chairs",
                                Id =drsql1["AccountCode"].ToString()//"101"
                            },

                    }
                    };
                }
            }
            return inlineDefault;
        }
        public ActionResult CheckList_Read([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM GLChartOfAccounts /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                          AccountID, AccountCode, AccountName, Prefix, (select top 1 CountryName from ListCountries where ListCountries.CountryISONumericCode=GLChartOfAccounts.CountryISONumericCode ) as  CountryISONumericCode, CurrencyISOCode, ClassID,AccountStatus FROM GLChartOfAccounts
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";
            //     const string selectQuery = @"SELECT
            //CheckListId, CheckTypeId, (select CheckType  from CheckTypes where CheckTypes.CheckTypeId = CheckList.CheckTypeId ) as CheckTypeName, CheckText, case when Status = 'true' then 'Active' else 'Inactive' end as StatusName, Status FROM CheckList";
            SqlBuilder builder = new SqlBuilder();
            var count = builder.AddTemplate(countQuery);

            int CurrentPage = 0; int CurrentPageSize = 1;
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

            builder.Where("isHead='true'");

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("AccountID");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);

            var rows = _dbcontext.Query<ChartOfAccountsModel>(selector.RawSql, selector.Parameters);

            //  rows.Each(item => Console.WriteLine($"({ item.AccountCode}): { item.AccountName}
            //{                GetParentsString(items, item)}"));            
            //}
            foreach (ChartOfAccountsModel model in rows)
            {
                string Path = GetParentsString(model.AccountCode);
                if (Path.Trim() != "")
                {
                    Path = Path.Substring(2);
                }
                model.ParentAccountName = Path;
            }

            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }
        static string GetParentsString(string pAccountCode)
        {
            string path = "";
            if (pAccountCode.Trim() != "")
            {
                DataTable dtAccount = DBUtils.GetDataTable("select AccountName, ParentAccountCode from  GLChartOfAccounts where AccountCode='" + pAccountCode + "'");
                if (dtAccount != null)
                {
                    if (dtAccount.Rows.Count > 0)
                    {
                        path = GetParentsString(Common.toString(dtAccount.Rows[0]["ParentAccountCode"])) + "->" + Common.toString(dtAccount.Rows[0]["AccountName"]) + path;
                    }
                }
            }
            return path;
        }

        //static string GetParentsString(List<ChartOfAccountsModel> all, ChartOfAccountsModel current)
        //{
        //    string path = "";
        //    Action<List<ChartOfAccountsModel>, ChartOfAccountsModel> GetPath = null;
        //    GetPath = (List<ChartOfAccountsModel> ps, ChartOfAccountsModel p) =>
        //    {
        //        var parents = all.Where(x => x.AccountCode == p.ParentAccountCode);
        //        foreach (var parent in parents)
        //        {
        //            path += $"/{ parent.AccountName}";
        //            GetPath(ps, parent);
        //        }
        //    };
        //    GetPath(all, current);
        //    string[] split = path.Split(new char[] { '/' });
        //    Array.Reverse(split);
        //    return string.Join("/", split);
        //}

        public ActionResult GetMainParent(string pAccountCode)
        {
            string mainParent = "";
            bool isValid = true;
            mainParent = GetParent(pAccountCode);
            if (mainParent == "4" || mainParent == "5")
            {
                isValid = false;
            }
            return Json(isValid, JsonRequestBehavior.AllowGet);
        }
        public static string GetParent(string pAccountCode)
        {
            string mainParent = "";
            if (pAccountCode.Trim() != "")
            {
                DataTable dtAccount = DBUtils.GetDataTable("select AccountName, ParentAccountCode from  GLChartOfAccounts where AccountCode='" + pAccountCode + "'");
                if (dtAccount != null)
                {
                    if (dtAccount.Rows.Count > 0)
                    {
                        mainParent = Common.toString(dtAccount.Rows[0]["ParentAccountCode"]);
                        if (mainParent != "0")
                        {
                            mainParent = GetParent(mainParent);
                        }
                        else
                        {
                            mainParent = pAccountCode;
                        }
                    }
                }
            }
            return mainParent;
        }
        public ActionResult hCreate(string Id)
        {
            ChartOfAccountsModel pModel = new ChartOfAccountsModel();
            int myId = Common.toInt(Id);

            #region Entities Code
            // var getlist = db.GLChartOfAccounts.Where(x => x.AccountID == myId).FirstOrDefault();
            //  if (getlist != null)
            //{
            //    pModel.ParentAccountName = getlist.AccountName;
            //    pModel.ParentAccountCode = Common.toString(getlist.AccountCode);
            //}
            #endregion

            ////string sql = "select AccountCode,AccountName from GLChartOfAccounts where AccountID=" + myId + "";
            ////DataTable dt = DBUtils.GetDataTable(sql);
            ////if (dt != null && dt.Rows.Count > 0)
            ////{
            ////    pModel.ParentAccountCode = Common.toString(dt.Rows[0]["AccountCode"]);
            ////    pModel.ParentAccountName = Common.toString(dt.Rows[0]["AccountName"]);
            ////}


            string sql = "select AccountCode,AccountName from GLChartOfAccounts where AccountID=" + myId + "";
            DataTable getlist = DBUtils.GetDataTable(sql);
            if (getlist != null)
            {
                if (getlist.Rows.Count > 0)
                {
                    DataRow drChartofAccounts = getlist.Rows[0];
                    string getAccountCount = Common.toString(DBUtils.executeSqlGetSingle("select Count(*) from GLChartOfAccounts where ParentAccountCode=" + drChartofAccounts["AccountCode"].ToString() + ""));
                    pModel.ParentAccountName = Common.toString(drChartofAccounts["AccountName"]); //getlist.AccountName;
                    pModel.ParentAccountCode = Common.toString(drChartofAccounts["AccountCode"]); // Common.toString(getlist.AccountCode);

                    //if (getAccountCount == "")
                    //{
                    //    getAccountCount = "0";
                    //}
                    //if (Common.toInt(getAccountCount) <= 9)
                    //{
                    //    pModel.ParentAccountName = Common.toString(drChartofAccounts["AccountName"]); //getlist.AccountName;
                    //    pModel.ParentAccountCode = Common.toString(drChartofAccounts["AccountCode"]); // Common.toString(getlist.AccountCode);

                    //}
                    //else
                    //{
                    //    pModel.isVisible = true;
                    //    TempData["Failed"] = "Already 10 subAccount exist for this Account";
                    //}
                }
            }

            pModel.AccountStatus = true;
            DropDownTreeItemModel inlineDefault = new DropDownTreeItemModel();
            ViewBag.inlineDefault = GetChartOfAccountClassicification();

            return View("~/Views/ChartOfAccounts/hCreate.cshtml", pModel);
        }

        [HttpPost]
        public ActionResult hCreate(ChartOfAccountsModel chartOfAccounts)
        {
            DropDownTreeItemModel inlineDefault = new DropDownTreeItemModel();
            string mainParent = "";
            ViewBag.inlineDefault = GetChartOfAccountClassicification();
            if (ModelState.IsValid)
            {
                if (Common.toString(chartOfAccounts.AccountName) != "" && Common.toString(chartOfAccounts.ParentAccountCode) != "" && Common.toString(chartOfAccounts.CountryISONumericCode) != "" && Common.toString(chartOfAccounts.CurrencyISOCode) != "" && Common.toString(chartOfAccounts.Prefix) != "")
                {
                    if (chartOfAccounts.AccountName == "Income" || chartOfAccounts.AccountName == "Expense")
                    {
                        chartOfAccounts.RevalueYN = false;
                        chartOfAccounts.Revalue = 0;
                    }
                    else
                    {
                        mainParent = GetParent(Common.toString(chartOfAccounts.ParentAccountCode));
                        if (mainParent == "4" || mainParent == "5")
                        {
                            chartOfAccounts.RevalueYN = false;
                            chartOfAccounts.Revalue = 0;
                        }
                    }
                    bool isUnique = false;
                    string LastChildAccountCode = "";
                    //string SqlGetlastChildAccountCode = "select top 1 AccountCode  from GLChartOfAccounts where ParentAccountCode = '" + chartOfAccounts.ParentAccountCode + "' order by AccountID desc";
                    //LastChildAccountCode = DBUtils.executeSqlGetSingle(SqlGetlastChildAccountCode);
                    //if (!string.IsNullOrEmpty(LastChildAccountCode))
                    //{
                    //    LastChildAccountCode = (1 + Common.toInt(LastChildAccountCode)).ToString();
                    //    LastChildAccountCode = LastChildAccountCode.PadLeft(4, '0');
                    //}
                    //else
                    //{
                    //    LastChildAccountCode = chartOfAccounts.ParentAccountCode + "1".PadLeft(4, '0');
                    //}
                    string getAccountCount = Common.toString(DBUtils.executeSqlGetSingle("select Count(*) from GLChartOfAccounts where ParentAccountCode='" + chartOfAccounts.ParentAccountCode + "'"));
                    if (getAccountCount == "")
                    {
                        getAccountCount = "0";
                    }
                    //  if (Common.toInt(getAccountCount) <= 9)
                    // {
                    for (int i = 1; i <= 500; i++)
                    {
                        if (chartOfAccounts.ParentAccountCode.Length == 4)
                        {
                            LastChildAccountCode = chartOfAccounts.ParentAccountCode + Common.toString(i).PadLeft(4, '0');
                        }
                        else
                        {
                            LastChildAccountCode = chartOfAccounts.ParentAccountCode + Common.toString(i);
                        }

                        string sql = "select AccountCode  from GLChartOfAccounts where  AccountCode = '" + LastChildAccountCode + "' and ParentAccountCode = '" + chartOfAccounts.ParentAccountCode + "' order by AccountCode asc";
                        string pCode = Common.toString(DBUtils.executeSqlGetSingle(sql));
                        if (pCode == "")
                        {
                            // code1 = Common.toString(Common.toInt(code1) + 1);
                            chartOfAccounts.AccountCode = LastChildAccountCode;
                            break;
                        }
                    }
                    ////string AccountName = PostingUtils.getGLAccountNameByCode(LastChildAccountCode);
                    ////if (Common.toString(AccountName).Trim() != "")
                    ////{
                    ////    int seed = 1;
                    ////    while (isUnique == false)
                    ////    {
                    ////        if (!string.IsNullOrEmpty(LastChildAccountCode))
                    ////        {
                    ////            LastChildAccountCode = (seed + Common.toInt(LastChildAccountCode)).ToString();
                    ////            LastChildAccountCode = LastChildAccountCode.PadLeft(4, '0');
                    ////        }
                    ////        else
                    ////        {
                    ////            LastChildAccountCode = chartOfAccounts.ParentAccountCode + "1".PadLeft(4, '0');
                    ////        }
                    ////        AccountName = PostingUtils.getGLAccountNameByCode(LastChildAccountCode);
                    ////        if (string.IsNullOrEmpty(AccountName))
                    ////        {
                    ////            isUnique = true;
                    ////        }
                    ////        else
                    ////        {
                    ////            isUnique = false;
                    ////        }
                    ////        seed++;
                    ////    }
                    ////}
                    chartOfAccounts.AccountCode = LastChildAccountCode;
                    string SqlInsert = "Insert Into GLChartOfAccounts (AccountCode,AccountName,ParentAccountCode, CountryISONumericCode,CurrencyISOCode,Prefix,Revalue,RevalueYN,DebiteLimit,CreditLimit,AccountStatus,isHead, AddedBy, AddedDate)";
                    SqlInsert += "Values(@AccountCode,@AccountName,@ParentAccountCode, @CountryISONumericCode,@CurrencyISOCode,@Prefix,@Revalue,@RevalueYN,@DebiteLimit,@CreditLimit,@AccountStatus,@isHead, @AddedBy, @AddedDate)";
                    SqlConnection conn = Common.getConnection();
                    ////using (SqlConnection conn = Common.getConnection())
                    ////{
                    SqlCommand insert = conn.CreateCommand();
                    try
                    {
                        insert.Parameters.AddWithValue("@AccountCode", chartOfAccounts.AccountCode);
                        insert.Parameters.AddWithValue("@AccountName", chartOfAccounts.AccountName);
                        insert.Parameters.AddWithValue("@ParentAccountCode", chartOfAccounts.ParentAccountCode);
                        insert.Parameters.AddWithValue("@CountryISONumericCode", chartOfAccounts.CountryISONumericCode);
                        insert.Parameters.AddWithValue("@CurrencyISOCode", chartOfAccounts.CurrencyISOCode);
                        insert.Parameters.AddWithValue("@Prefix", chartOfAccounts.Prefix);
                        insert.Parameters.AddWithValue("@Revalue", chartOfAccounts.Revalue);
                        insert.Parameters.AddWithValue("@RevalueYN", chartOfAccounts.RevalueYN);
                        insert.Parameters.AddWithValue("@DebiteLimit", chartOfAccounts.DebiteLimit);
                        insert.Parameters.AddWithValue("@CreditLimit", chartOfAccounts.CreditLimit);
                        insert.Parameters.AddWithValue("@AddedBy", Profile.Id);
                        insert.Parameters.AddWithValue("@AddedDate", DateTime.Now);
                        insert.Parameters.AddWithValue("@AccountStatus", chartOfAccounts.AccountStatus);
                        insert.Parameters.AddWithValue("@isHead", 1);
                        insert.CommandText = SqlInsert;
                        int RowsAffected = insert.ExecuteNonQuery();
                        if (RowsAffected == 1)
                        {
                            TempData["Success"] = "Chart of account heads successfully created";
                            return View("~/Views/ChartOfAccounts/hCreate.cshtml", chartOfAccounts);
                        }
                        else
                        {
                            TempData["Failed"] = "Error occured: ";
                            return View("~/Views/ChartOfAccounts/hCreate.cshtml", chartOfAccounts);
                        }
                    }
                    catch (Exception ex)
                    {
                        TempData["Failed"] = "Error occured: " + ex.Message.ToString();
                        return View("~/Views/ChartOfAccounts/hCreate.cshtml", chartOfAccounts);
                    }
                    finally
                    {
                        conn.Close();
                    }
                    //// }
                    // }
                    //else
                    //{
                    //    TempData["Failed"] = "Already 10 subAccount exist for this Account.";
                    //    return View("~/Views/ChartOfAccounts/hCreate.cshtml", chartOfAccounts);
                    //}
                }
                else
                {
                    TempData["Failed"] = "Data missing. Account Classification, Name, Prefix, Country and Currency are required.";
                    return View("~/Views/ChartOfAccounts/hCreate.cshtml", chartOfAccounts);
                }
            }
            else
            {
                TempData["Failed"] = "Data missing. Please check all fields and try again";
                return View("~/Views/ChartOfAccounts/hCreate.cshtml", chartOfAccounts);
            }
        }
        public ActionResult hChangeCategory(string id)
        {
            DropDownTreeItemModel inlineDefault = new DropDownTreeItemModel();
            ViewBag.inlineDefault = GetChartOfAccountClassicification();
            ChartOfAccountsModel accountsModel = new ChartOfAccountsModel();
            if (!string.IsNullOrEmpty(id))
            {
                int pId = Common.toInt(id);
                // var accounts = db.GLChartOfAccounts.Where(x => x.AccountID == pId).Take(1).FirstOrDefault();
                //accountsModel.preAccountID = accounts.AccountID;
                //accountsModel.preAccountCode = accounts.AccountCode;
                //accountsModel.AccountCodeMask = accounts.AccountCodeMask;
                //accountsModel.AccountDescription = accounts.AccountDescription;
                //accountsModel.AccountName = accounts.AccountName;
                //accountsModel.preParentAccountCode = accounts.ParentAccountCode;
                //accountsModel.AccountStatus = Common.toBool(accounts.AccountStatus);
                //accountsModel.Charges = Common.toBool(accounts.Charges);
                //accountsModel.ClassID = Common.toInt(accounts.ClassID);
                //accountsModel.CountryISONumericCode = accounts.CountryISONumericCode;
                //accountsModel.CreditLimit = Common.toDecimal(accounts.CreditLimit);
                //accountsModel.CurrencyISOCode = accounts.CurrencyISOCode;
                //accountsModel.DebiteLimit = Common.toDecimal(accounts.DebiteLimit);
                //accountsModel.Prefix = accounts.Prefix;
                //accountsModel.Revalue = Common.toDecimal(accounts.Revalue);
                //accountsModel.RevalueYN = Common.toBool(accounts.RevalueYN);
                //accountsModel.TaxPercentage = Common.toDecimal(accounts.TaxPercentage);
                //accountsModel.TypeID = Common.toInt(accounts.TypeID);
                //accountsModel.VATWHT = Common.toDecimal(accounts.VATWHT);
                //accountsModel.AddedBy = accounts.AddedBy;
                //accountsModel.AddedDate = Common.toDateTime(accounts.AddedDate);
                string sql = "select * from GLChartOfAccounts where AccountID=" + pId + "";
                DataTable dt = DBUtils.GetDataTable(sql);
                if (dt != null && dt.Rows.Count > 0)
                {
                    accountsModel.preAccountID = Common.toInt(dt.Rows[0]["AccountID"]);
                    accountsModel.preAccountCode = Common.toString(dt.Rows[0]["AccountCode"]);
                    accountsModel.AccountCodeMask = Common.toString(dt.Rows[0]["AccountCodeMask"]);
                    accountsModel.AccountDescription = Common.toString(dt.Rows[0]["AccountDescription"]);
                    accountsModel.AccountName = Common.toString(dt.Rows[0]["AccountName"]);
                    accountsModel.preParentAccountCode = Common.toString(dt.Rows[0]["ParentAccountCode"]);
                    accountsModel.AccountStatus = Common.toBool(dt.Rows[0]["AccountStatus"]);
                    accountsModel.Charges = Common.toBool(dt.Rows[0]["Charges"]);
                    accountsModel.ClassID = Common.toInt(dt.Rows[0]["ClassID"]);
                    accountsModel.CountryISONumericCode = Common.toString(dt.Rows[0]["CountryISONumericCode"]);
                    accountsModel.CreditLimit = Common.toDecimal(dt.Rows[0]["CreditLimit"]);
                    accountsModel.CurrencyISOCode = Common.toString(dt.Rows[0]["CurrencyISOCode"]);
                    accountsModel.DebiteLimit = Common.toDecimal(dt.Rows[0]["DebiteLimit"]);
                    accountsModel.Prefix = Common.toString(dt.Rows[0]["Prefix"]);
                    accountsModel.Revalue = Common.toDecimal(dt.Rows[0]["Revalue"]);
                    accountsModel.RevalueYN = Common.toBool(dt.Rows[0]["RevalueYN"]);
                    accountsModel.TaxPercentage = Common.toDecimal(dt.Rows[0]["TaxPercentage"]);
                    accountsModel.TypeID = Common.toInt(dt.Rows[0]["TypeID"]);
                    accountsModel.VATWHT = Common.toDecimal(dt.Rows[0]["VATWHT"]);
                    accountsModel.AddedBy = Common.toString(dt.Rows[0]["AddedBy"]);
                    accountsModel.AddedDate = Common.toDateTime(dt.Rows[0]["AddedDate"]);
                }
                string Path = GetParentsString(accountsModel.preAccountCode);

                if (Path.Trim() != "")
                {
                    Path = Path.Substring(2);
                }
                accountsModel.preParentAccountName = Path;

                return View("~/Views/ChartOfAccounts/hChangeCategory.cshtml", accountsModel);
            }
            else
            {
                TempData["Failed"] = "Invalid AccountId.";
                return View("~/Views/ChartOfAccounts/hChangeCategory.cshtml", accountsModel);

            }
        }
        [HttpPost]
        public ActionResult hChangeCategory(ChartOfAccountsModel chartOfAccounts)
        {
            DropDownTreeItemModel inlineDefault = new DropDownTreeItemModel();
            ViewBag.inlineDefault = GetChartOfAccountClassicification();
            string pCode = "";
            if (!string.IsNullOrEmpty(chartOfAccounts.PinCode))
            {
                bool isValid = Common.ValidatePinCode(Profile.Id, chartOfAccounts.PinCode);
                if (isValid)
                {
                    if (!string.IsNullOrEmpty(chartOfAccounts.preAccountCode) && chartOfAccounts.preAccountID > 0 && !string.IsNullOrEmpty(chartOfAccounts.ParentAccountCode))
                    {
                        //string pCode = DBUtils.executeSqlGetSingle("select AccountCode from GLChartOfAccounts where ParentAccountCode='" + chartOfAccounts.ParentAccountCode + "'");
                        ////if (string.IsNullOrEmpty(pCode))
                        //{
                        if (ModelState.IsValid)
                        {
                            if (Common.toString(chartOfAccounts.ParentAccountCode) != "")
                            {
                                bool isUnique = false; int i = 1;
                                string LastChildAccountCode = "";
                                string SqlGetlastChildAccountCode = "select top 1 AccountCode  from GLChartOfAccounts where ParentAccountCode = '" + chartOfAccounts.ParentAccountCode + "' order by AccountID desc";
                                LastChildAccountCode = DBUtils.executeSqlGetSingle(SqlGetlastChildAccountCode);
                                if (!string.IsNullOrEmpty(LastChildAccountCode))
                                {
                                    while (isUnique == false)
                                    {
                                        LastChildAccountCode = chartOfAccounts.ParentAccountCode + Common.toString(i).PadLeft(4, '0');
                                        string sql = "select AccountCode  from GLChartOfAccounts where  AccountCode = '" + LastChildAccountCode + "' and ParentAccountCode = '" + chartOfAccounts.ParentAccountCode + "' order by AccountCode asc";
                                        pCode = DBUtils.executeSqlGetSingle(sql);
                                        if (!string.IsNullOrEmpty(pCode))
                                        {
                                            i++;
                                        }
                                        else
                                        {
                                            isUnique = true;
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    LastChildAccountCode = chartOfAccounts.ParentAccountCode + "1".PadLeft(4, '0');
                                }
                                chartOfAccounts.AccountCode = LastChildAccountCode;

                                string UpdateSQL = "Update GLChartOfAccounts Set AccountCode='" + chartOfAccounts.AccountCode + "',  ParentAccountCode= '" + chartOfAccounts.ParentAccountCode + "'";
                                UpdateSQL += " Where AccountID =" + chartOfAccounts.preAccountID + "";
                                bool OK = DBUtils.ExecuteSQL(UpdateSQL);
                                if (OK)
                                {
                                    DBUtils.ExecuteSQL("update GLTransactions set AccountCode='" + chartOfAccounts.AccountCode + "' where AccountCode='" + chartOfAccounts.preAccountCode + "'");
                                    DBUtils.ExecuteSQL("update Customers set AccountCode='" + chartOfAccounts.AccountCode + "' where AccountCode='" + chartOfAccounts.preAccountCode + "'");
                                    DBUtils.ExecuteSQL("update Vendors set AccountCode='" + chartOfAccounts.AccountCode + "' where AccountCode='" + chartOfAccounts.preAccountCode + "'");
                                    DBUtils.ExecuteSQL("update VendorBillDetails set BillAccountCode='" + chartOfAccounts.AccountCode + "' where BillAccountCode='" + chartOfAccounts.preAccountCode + "'");
                                    DBUtils.ExecuteSQL("update VendorBillPayment set AccountCode='" + chartOfAccounts.AccountCode + "' where AccountCode='" + chartOfAccounts.preAccountCode + "'");
                                    DBUtils.ExecuteSQL("update CustomerInvoices set AccountCode='" + chartOfAccounts.AccountCode + "' where AccountCode='" + chartOfAccounts.preAccountCode + "'");
                                    DBUtils.ExecuteSQL("update GLBankAccounts set AccountCode='" + chartOfAccounts.AccountCode + "' where AccountCode='" + chartOfAccounts.preAccountCode + "'");

                                    //DBUtils.ExecuteSQL("update CustomerTransactions set AgentAccountCode='" + chartOfAccounts.AccountCode + "' where AgentAccountCode='" + chartOfAccounts.preAccountCode + "'");
                                    //DBUtils.ExecuteSQL("update CustomerTransactions set HoldAccountCode='" + chartOfAccounts.AccountCode + "' where HoldAccountCode='" + chartOfAccounts.preAccountCode + "'");
                                    //DBUtils.ExecuteSQL("update CustomerTransactions set BuyerAccountCode='" + chartOfAccounts.AccountCode + "' where BuyerAccountCode='" + chartOfAccounts.preAccountCode + "'");
                                    //DBUtils.ExecuteSQL("update CustomerTransactions set PaidAccountCode='" + chartOfAccounts.AccountCode + "' where PaidAccountCode='" + chartOfAccounts.preAccountCode + "'");

                                    TempData["Success"] = "Chart of account heads successfully created";
                                    return View("~/Views/ChartOfAccounts/hCreate.cshtml", chartOfAccounts);
                                }
                                else
                                {
                                    TempData["Failed"] = "Error occured: ";
                                    return View("~/Views/ChartOfAccounts/hChangeCategory.cshtml", chartOfAccounts);
                                }

                            }
                            else
                            {
                                TempData["Failed"] = "Data missing. Account Classification, Name, Prefix, Country and Currency are required.";
                                return View("~/Views/ChartOfAccounts/hChangeCategory.cshtml", chartOfAccounts);
                            }
                        }
                        else
                        {
                            TempData["Failed"] = "Data missing. Please check all fields and try again";
                            return View("~/Views/ChartOfAccounts/hChangeCategory.cshtml", chartOfAccounts);
                        }
                    }
                    else
                    {
                        TempData["Failed"] = "Data missing. Please check all fields and try again";
                        return View("~/Views/ChartOfAccounts/hChangeCategory.cshtml", chartOfAccounts);
                    }
                }
                else
                {
                    TempData["Failed"] = "Invalid pin code.Please provide valid pin code.";
                    return View("~/Views/ChartOfAccounts/hChangeCategory.cshtml", chartOfAccounts);
                }
            }
            else
            {
                TempData["Failed"] = "Please provide valid pin code.";
                return View("~/Views/ChartOfAccounts/hChangeCategory.cshtml", chartOfAccounts);
            }
        }
        public ActionResult hEdit(int Id)
        {
            //var accounts = db.GLChartOfAccounts.Where(x => x.AccountID == Id).Take(1).FirstOrDefault();
            //ChartOfAccountsModel accountsModel = new ChartOfAccountsModel();
            //accountsModel.AccountID = accounts.AccountID;
            //accountsModel.AccountCode = accounts.AccountCode;
            //accountsModel.AccountCodeMask = accounts.AccountCodeMask;
            //accountsModel.AccountDescription = accounts.AccountDescription;
            //accountsModel.AccountName = accounts.AccountName;
            //accountsModel.AccountStatus = Common.toBool(accounts.AccountStatus);
            //accountsModel.Charges = Common.toBool(accounts.Charges);
            //accountsModel.ClassID = Common.toInt(accounts.ClassID);
            //accountsModel.CountryISONumericCode = accounts.CountryISONumericCode;
            //accountsModel.CreditLimit = Common.toDecimal(accounts.CreditLimit);
            //accountsModel.CurrencyISOCode = accounts.CurrencyISOCode;
            //accountsModel.DebiteLimit = Common.toDecimal(accounts.DebiteLimit);
            //accountsModel.ParentAccountCode = accounts.ParentAccountCode;
            //accountsModel.Prefix = accounts.Prefix;
            //accountsModel.Revalue = Common.toDecimal(accounts.Revalue);
            //accountsModel.RevalueYN = Common.toBool(accounts.RevalueYN);
            //accountsModel.TaxPercentage = Common.toDecimal(accounts.TaxPercentage);
            //accountsModel.TypeID = Common.toInt(accounts.TypeID);
            //accountsModel.VATWHT = Common.toDecimal(accounts.VATWHT);
            //accountsModel.AddedBy = accounts.AddedBy;
            //accountsModel.AddedDate = Common.toDateTime(accounts.AddedDate);
            ChartOfAccountsModel accountsModel = new ChartOfAccountsModel();
            string sql = "select * from GLChartOfAccounts where AccountID=" + Id + "";
            DataTable dt = DBUtils.GetDataTable(sql);
            if (dt != null && dt.Rows.Count > 0)
            {
                accountsModel.AccountID = Common.toInt(dt.Rows[0]["AccountID"]);
                accountsModel.AccountCode = Common.toString(dt.Rows[0]["AccountCode"]);
                accountsModel.AccountCodeMask = Common.toString(dt.Rows[0]["AccountCodeMask"]);
                accountsModel.AccountDescription = Common.toString(dt.Rows[0]["AccountDescription"]);
                accountsModel.AccountName = Common.toString(dt.Rows[0]["AccountName"]);
                accountsModel.ParentAccountCode = Common.toString(dt.Rows[0]["ParentAccountCode"]);
                accountsModel.AccountStatus = Common.toBool(dt.Rows[0]["AccountStatus"]);
                accountsModel.Charges = Common.toBool(dt.Rows[0]["Charges"]);
                accountsModel.ClassID = Common.toInt(dt.Rows[0]["ClassID"]);
                accountsModel.CountryISONumericCode = Common.toString(dt.Rows[0]["CountryISONumericCode"]);
                accountsModel.CreditLimit = Common.toDecimal(dt.Rows[0]["CreditLimit"]);
                accountsModel.CurrencyISOCode = Common.toString(dt.Rows[0]["CurrencyISOCode"]);
                accountsModel.DebiteLimit = Common.toDecimal(dt.Rows[0]["DebiteLimit"]);
                accountsModel.Prefix = Common.toString(dt.Rows[0]["Prefix"]);
                accountsModel.Revalue = Common.toDecimal(dt.Rows[0]["Revalue"]);
                accountsModel.RevalueYN = Common.toBool(dt.Rows[0]["RevalueYN"]);
                accountsModel.TaxPercentage = Common.toDecimal(dt.Rows[0]["TaxPercentage"]);
                accountsModel.TypeID = Common.toInt(dt.Rows[0]["TypeID"]);
                accountsModel.VATWHT = Common.toDecimal(dt.Rows[0]["VATWHT"]);
                accountsModel.AddedBy = Common.toString(dt.Rows[0]["AddedBy"]);
                accountsModel.AddedDate = Common.toDateTime(dt.Rows[0]["AddedDate"]);
            }
            return View("~/Views/ChartOfAccounts/hEdit.cshtml", accountsModel);
        }

        [HttpPost]
        public ActionResult hEdit(ChartOfAccountsModel chartOfAccounts)
        {
            if (ModelState.IsValid)
            {
                string mainParent = "";
                if (chartOfAccounts.AccountName == "Income" || chartOfAccounts.AccountName == "Expense")
                {
                    chartOfAccounts.RevalueYN = false;
                    chartOfAccounts.Revalue = 0;
                }
                else
                {
                    mainParent = GetParent(Common.toString(chartOfAccounts.AccountCode));
                    if (mainParent == "4" || mainParent == "5")
                    {
                        chartOfAccounts.RevalueYN = false;
                        chartOfAccounts.Revalue = 0;
                    }
                }
                string UpdateSQL = "Update GLChartOfAccounts Set AccountCode=@AccountCode, AccountName= @AccountName,";
                UpdateSQL += "CountryISONumericCode=@CountryISONumericCode,CurrencyISOCode=@CurrencyISOCode,Prefix=@Prefix,Revalue=@Revalue,RevalueYN=@RevalueYN,DebiteLimit=@DebiteLimit,CreditLimit=@CreditLimit,AccountStatus=@AccountStatus Where AccountID =@AccountID";

                SqlConnection con = Common.getConnection();
                ////using (SqlConnection con = Common.getConnection())
                ////{
                SqlCommand updateCommand = con.CreateCommand();

                try
                {
                    updateCommand.Parameters.AddWithValue("@AccountCode", chartOfAccounts.AccountCode);
                    updateCommand.Parameters.AddWithValue("@AccountName", chartOfAccounts.AccountName);
                    updateCommand.Parameters.AddWithValue("@CountryISONumericCode", chartOfAccounts.CountryISONumericCode);
                    updateCommand.Parameters.AddWithValue("@CurrencyISOCode", chartOfAccounts.CurrencyISOCode);
                    updateCommand.Parameters.AddWithValue("@Prefix", chartOfAccounts.Prefix);
                    updateCommand.Parameters.AddWithValue("@Revalue", chartOfAccounts.Revalue);
                    updateCommand.Parameters.AddWithValue("@RevalueYN", chartOfAccounts.RevalueYN);
                    updateCommand.Parameters.AddWithValue("@DebiteLimit", chartOfAccounts.DebiteLimit);
                    updateCommand.Parameters.AddWithValue("@CreditLimit", chartOfAccounts.CreditLimit);
                    //updateCommand.Parameters.AddWithValue("@AddedBy", chartOfAccounts.AddedBy);
                    //updateCommand.Parameters.AddWithValue("@AddedDate", chartOfAccounts.AddedDate);
                    updateCommand.Parameters.AddWithValue("@AccountStatus", chartOfAccounts.AccountStatus);
                    updateCommand.Parameters.AddWithValue("@AccountID", chartOfAccounts.AccountID);
                    updateCommand.CommandText = UpdateSQL;
                    int RowsAffected = updateCommand.ExecuteNonQuery();
                    if (RowsAffected == 1)
                    {
                        TempData["Success"] = "Chart of account heads updated successfully";
                        return View("~/Views/ChartOfAccounts/hEdit.cshtml", chartOfAccounts);
                    }
                    else
                    {
                        TempData["Failed"] = "Error occured: ";
                        return View("~/Views/ChartOfAccounts/hEdit.cshtml", chartOfAccounts);
                    }
                    //  con.Close();
                }
                catch (Exception ex)
                {
                    TempData["Failed"] = "Error occured: " + ex.Message.ToString();
                    return View("~/Views/ChartOfAccounts/hEdit.cshtml", chartOfAccounts);
                }

                //// }

            }
            else
            {
                TempData["Failed"] = "Data missing. Please check all fields and try again";
                return View("~/Views/ChartOfAccounts/hEdit.cshtml", chartOfAccounts);

            }
        }
        public ActionResult hDelete(int Id)
        {
            //var accounts = db.GLChartOfAccounts.Find(Id);
            //ChartOfAccountsModel accountsModel = new ChartOfAccountsModel();
            //accountsModel.AccountID = accounts.AccountID;
            //accountsModel.AccountCode = accounts.AccountCode;
            //accountsModel.AccountCodeMask = accounts.AccountCodeMask;
            //accountsModel.AccountDescription = accounts.AccountDescription;
            //accountsModel.AccountName = accounts.AccountName;
            //accountsModel.AccountStatus = Common.toBool(accounts.AccountStatus);
            //accountsModel.Charges = Common.toBool(accounts.Charges);
            //accountsModel.ClassID = Common.toInt(accounts.ClassID);
            //accountsModel.CountryISONumericCode = Common.toString(accounts.CountryISONumericCode);
            //accountsModel.CreditLimit = Common.toDecimal(accounts.CreditLimit);
            //accountsModel.CurrencyISOCode = accounts.CurrencyISOCode;
            //accountsModel.DebiteLimit = Common.toDecimal(accounts.DebiteLimit);
            //accountsModel.ParentAccountCode = accounts.ParentAccountCode;
            //accountsModel.Prefix = accounts.Prefix;
            //accountsModel.Revalue = Common.toDecimal(accounts.Revalue);
            //accountsModel.RevalueYN = Common.toBool(accounts.RevalueYN);
            //accountsModel.TaxPercentage = Common.toDecimal(accounts.TaxPercentage);
            //accountsModel.TypeID = Common.toInt(accounts.TypeID);
            //accountsModel.VATWHT = Common.toDecimal(accounts.VATWHT);
            //accountsModel.AddedBy = accounts.AddedBy;
            //accountsModel.AddedDate = Common.toDateTime(accounts.AddedDate);
            ChartOfAccountsModel accountsModel = new ChartOfAccountsModel();
            string sql = "select * from GLChartOfAccounts where AccountID=" + Id + "";
            DataTable dt = DBUtils.GetDataTable(sql);
            if (dt != null && dt.Rows.Count > 0)
            {
                accountsModel.AccountID = Common.toInt(dt.Rows[0]["AccountID"]);
                accountsModel.AccountCode = Common.toString(dt.Rows[0]["AccountCode"]);
                accountsModel.AccountCodeMask = Common.toString(dt.Rows[0]["AccountCodeMask"]);
                accountsModel.AccountDescription = Common.toString(dt.Rows[0]["AccountDescription"]);
                accountsModel.AccountName = Common.toString(dt.Rows[0]["AccountName"]);
                accountsModel.ParentAccountCode = Common.toString(dt.Rows[0]["ParentAccountCode"]);
                accountsModel.AccountStatus = Common.toBool(dt.Rows[0]["AccountStatus"]);
                accountsModel.Charges = Common.toBool(dt.Rows[0]["Charges"]);
                accountsModel.ClassID = Common.toInt(dt.Rows[0]["ClassID"]);
                accountsModel.CountryISONumericCode = Common.toString(dt.Rows[0]["CountryISONumericCode"]);
                accountsModel.CreditLimit = Common.toDecimal(dt.Rows[0]["CreditLimit"]);
                accountsModel.CurrencyISOCode = Common.toString(dt.Rows[0]["CurrencyISOCode"]);
                accountsModel.DebiteLimit = Common.toDecimal(dt.Rows[0]["DebiteLimit"]);
                accountsModel.Prefix = Common.toString(dt.Rows[0]["Prefix"]);
                accountsModel.Revalue = Common.toDecimal(dt.Rows[0]["Revalue"]);
                accountsModel.RevalueYN = Common.toBool(dt.Rows[0]["RevalueYN"]);
                accountsModel.TaxPercentage = Common.toDecimal(dt.Rows[0]["TaxPercentage"]);
                accountsModel.TypeID = Common.toInt(dt.Rows[0]["TypeID"]);
                accountsModel.VATWHT = Common.toDecimal(dt.Rows[0]["VATWHT"]);
                accountsModel.AddedBy = Common.toString(dt.Rows[0]["AddedBy"]);
                accountsModel.AddedDate = Common.toDateTime(dt.Rows[0]["AddedDate"]);
            }
            return View("~/Views/ChartOfAccounts/hDelete.cshtml", accountsModel);
        }
        [HttpPost]
        public ActionResult hDelete(ChartOfAccountsModel chartOfAccounts)
        {
            if (chartOfAccounts.AccountID > 0)
            {
                try
                {
                    //GLChartOfAccounts account = db.GLChartOfAccounts.Find(chartOfAccounts.AccountID);
                    // if (account != null)
                    string sqlGLChartOfAccounts = "select * from GLChartOfAccounts where AccountID=" + chartOfAccounts.AccountID + "";
                    DataTable dt = DBUtils.GetDataTable(sqlGLChartOfAccounts);
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        string transactionID = DBUtils.executeSqlGetSingle("select TransactionID from GLTransactions where AccountCode='" + dt.Rows[0]["AccountCode"].ToString() + "'");
                        if (string.IsNullOrEmpty(Common.toString(transactionID)))
                        {
                            string sql = "select top 1 AccountCode from GLChartOfAccounts where ParentAccountCode = '" + dt.Rows[0]["AccountCode"].ToString() + "' order by AccountID desc";
                            string IsChild = DBUtils.executeSqlGetSingle(sql);
                            if (string.IsNullOrEmpty(IsChild))
                            {
                                string sqlDelete = "Delete From GLChartOfAccounts Where AccountID = '" + chartOfAccounts.AccountID + "'";
                                bool Isfound = DBUtils.ExecuteSQL(sqlDelete);
                                if (Isfound)
                                {
                                    TempData["Success"] = "Successfully deleted.";
                                    return View("~/Views/ChartOfAccounts/hdelete.cshtml", chartOfAccounts);
                                }
                                else
                                {
                                    TempData["Failed"] = "Error occured: ";
                                    return View("~/Views/ChartOfAccounts/hdelete.cshtml", chartOfAccounts);
                                }
                            }
                            else
                            {
                                TempData["Failed"] = "Can not delete, sub account exists.";
                                return View("~/Views/ChartOfAccounts/hdelete.cshtml", chartOfAccounts);
                            }
                        }
                        else
                        {
                            TempData["Failed"] = "Can not delete, this account head have some transaction in the database.";
                            return View("~/Views/ChartOfAccounts/hdelete.cshtml", chartOfAccounts);
                        }
                    }
                    else
                    {
                        TempData["Failed"] = "Error occured: Account not found.";
                        return View("~/Views/ChartOfAccounts/hdelete.cshtml", chartOfAccounts);
                    }
                }
                catch (Exception ex)
                {
                    TempData["Failed"] = "Error occured: " + ex.Message.ToString();
                    return View("~/Views/ChartOfAccounts/hdelete.cshtml", chartOfAccounts);
                }
            }
            else
            {
                TempData["Failed"] = "Invalid account id.";
                return View("~/Views/ChartOfAccounts/hdelete.cshtml", chartOfAccounts);
            }
        }
        #endregion

        #region Chart of Account Balances
        public ActionResult bIndex(ChartOfAccountsModel model)
        {
            if (model != null)
            {
                if (model.FromDate != Common.toDateTime("01/01/0001 00:00:00"))
                {
                    UpdateAccountBalances2(model.FromDate);
                }
                else
                {
                    model.FromDate = Common.toDateTime(SysPrefs.PostingDate);
                    UpdateAccountBalances2(Common.toDateTime(SysPrefs.PostingDate));
                }
            }
            else
            {
                model.FromDate = Common.toDateTime(SysPrefs.PostingDate);
                UpdateAccountBalances2(Common.toDateTime(SysPrefs.PostingDate));
            }
            return View("~/Views/ChartOfAccounts/bIndex.cshtml", model);
        }

        public static void UpdateAccountBalances(DateTime pPostingDate)
        {
            string strSql = "";
            strSql = "Update GLChartOfAccounts  set ForeignCurrencyBalance=0, LocalCurrencyBalance=0";
            DBUtils.ExecuteSQL(strSql);

            strSql = "";
            strSql = @"SELECT AccountCode, isNull( SUM(GLTransactions.ForeignCurrencyAmount), 0) as FCBalance, isNull( SUM(GLTransactions.LocalCurrencyAmount), 0) as LCBalance 
                          From GLTransactions INNER JOIN GLReferences ON GLTransactions.GLReferenceID = GLReferences.GLReferenceID Where GLReferences.isPosted=1 ";
            if (pPostingDate != null)
            {
                strSql += " AND GLTransactions.AddedDate <='" + pPostingDate.ToString("yyyy-MM-dd") + " 23:59:00'";
            }
            strSql += " group by GLTransactions.AccountCode ";

            DataTable Balances = DBUtils.GetDataTable(strSql);
            if (Balances != null)
            {
                foreach (DataRow drAccount in Balances.Rows)
                {
                    string sqlUpdte = "Update GLChartOfAccounts  set ForeignCurrencyBalance='" + drAccount["FCBalance"] + "', LocalCurrencyBalance='" + drAccount["LCBalance"] + "' Where AccountCode='" + drAccount["AccountCode"] + "'";
                    DBUtils.ExecuteSQL(sqlUpdte);
                    string AccountTree = PostingUtils.GetAccountTree(Common.toString(drAccount["AccountCode"]));
                    if (AccountTree != "")
                    {
                        if (AccountTree.Length > 0)
                        {
                            AccountTree = AccountTree.Substring(0, AccountTree.Length - 1);
                        }
                        if (AccountTree != "")
                        {
                            sqlUpdte = "Update GLChartOfAccounts set ForeignCurrencyBalance= ForeignCurrencyBalance + (" + drAccount["FCBalance"] + "), LocalCurrencyBalance = LocalCurrencyBalance + (" + drAccount["LCBalance"] + ")  Where AccountCode in ('" + AccountTree.Replace(",", "','") + "')";
                            DBUtils.ExecuteSQL(sqlUpdte);
                        }
                    }
                }
            }
        }

        public static void UpdateAccountBalances2(DateTime pPostingDate)
        {
            DataTable Balances = PostingUtils.UpdateAndgetBalances(pPostingDate);
            if (Balances != null && Balances.Rows.Count > 0)
            {
                foreach (DataRow drAccount in Balances.Rows)
                {
                    string sqlUpdate = ";WITH UpdateParentAccountBalances AS ( select ParentAccountCode as pCode from GLChartOfAccounts where AccountCode='" + drAccount["AccountCode"].ToString() + "'";
                    sqlUpdate += " union all  select ParentAccountCode from GLChartOfAccounts e INNER JOIN UpdateParentAccountBalances o ON o.pCode = e.AccountCode)";
                    sqlUpdate += " UPDATE dbo.GLChartOfAccounts SET GLChartOfAccounts.ForeignCurrencyBalance = ForeignCurrencyBalance + (" + drAccount["FCBalance"] + "),";
                    sqlUpdate += " GLChartOfAccounts.LocalCurrencyBalance = LocalCurrencyBalance + (" + drAccount["LCBalance"] + ") FROM UpdateParentAccountBalances";
                    sqlUpdate += " WHERE GLChartOfAccounts.AccountCode = UpdateParentAccountBalances.pCode and pCode<>0";
                    DBUtils.ExecuteSQL(sqlUpdate);
                    #region old code
                    //string AccountTree = PostingUtils.GetAccountTree(Common.toString(drAccount["AccountCode"]));
                    //if (AccountTree != "")
                    //{
                    //    if (AccountTree.Length > 0)
                    //    {
                    //        AccountTree = AccountTree.Substring(0, AccountTree.Length - 1);
                    //    }
                    //    if (AccountTree != "")
                    //    {
                    //        string sqlUpdte = ";WITH UpdateBalances AS(SELECT ForeignCurrencyBalance,LocalCurrencyBalance FROM GLChartOfAccounts where AccountCode in ('" + AccountTree.Replace(",", "','") + "')) UPDATE UpdateBalances SET ForeignCurrencyBalance =ForeignCurrencyBalance + (" + drAccount["FCBalance"] + "),LocalCurrencyBalance=LocalCurrencyBalance + (" + drAccount["LCBalance"] + ")";
                    //        DBUtils.ExecuteSQL(sqlUpdte);
                    //    }
                    //}
                    #endregion
                }
            }
        }
        public JsonResult AccountsWithBalance_Read([DataSourceRequest] DataSourceRequest request)
        {
            #region old code
            // var result = db.GLChartOfAccounts.Where(x => x.AccountStatus == true).OrderBy(x => x.AccountCode).ToTreeDataSourceResult(request,
            //    e => e.AccountCode,
            //    e => e.ParentAccountCode,
            //    e => e
            //);
            #endregion
            var result = _dbcontext.Query<ChartOfAccountsModel>("select AccountID, AccountName, AccountCode, ParentAccountCode, Prefix, CurrencyISOCode, ForeignCurrencyBalance, LocalCurrencyBalance, AddedDate, isHead from GLChartOfAccounts where AccountStatus=1 order by AccountCode").ToTreeDataSourceResult(request,
                           e => e.AccountCode,
                           e => e.ParentAccountCode,
                           e => e
                       );
            return Json(result, JsonRequestBehavior.AllowGet);

        }
        //public void AccountsListExportToExcel(ChartOfAccountsModel model)
        //{
        //    try
        //    {
        //        XLWorkbook workbook = new XLWorkbook();
        //        var worksheet = workbook.Worksheets.Add("GL chart of account balances");
        //        worksheet.Cell("A1").Value = @SysPrefs.SiteName;
        //        worksheet.Cell("A1").Style.Font.Bold = true;
        //        worksheet.Cell("A2").Value = "GL chart of account balances";
        //        worksheet.Cell("A2").Style.Font.Bold = true;
        //        worksheet.Cell("E2").Value = "As on Dated : " + @SysPrefs.PostingDate + "";
        //        worksheet.Cell("E2").Style.Font.Bold = true;
        //        worksheet.Cell("A3").Value = "Code";
        //        worksheet.Cell("B3").Value = "Title of Account";
        //        worksheet.Cell("C3").Value = "Prefix";
        //        worksheet.Cell("D3").Value = "Currency";
        //        worksheet.Cell("E3").Value = "F/C Balance";
        //        worksheet.Cell("F3").Value = "L/C Balance";
        //        worksheet.Cell("G3").Value = "Last Update";

        //        worksheet.Cell("A3").Style.Fill.BackgroundColor = XLColor.PersianRed;
        //        worksheet.Cell("B3").Style.Fill.BackgroundColor = XLColor.PersianRed;
        //        worksheet.Cell("C3").Style.Fill.BackgroundColor = XLColor.PersianRed;
        //        worksheet.Cell("D3").Style.Fill.BackgroundColor = XLColor.PersianRed;
        //        worksheet.Cell("E3").Style.Fill.BackgroundColor = XLColor.PersianRed;
        //        worksheet.Cell("F3").Style.Fill.BackgroundColor = XLColor.PersianRed;
        //        worksheet.Cell("G3").Style.Fill.BackgroundColor = XLColor.PersianRed;

        //        worksheet.Cell("A3").Style.Font.FontColor = XLColor.Black;
        //        worksheet.Cell("B3").Style.Font.FontColor = XLColor.Black;
        //        worksheet.Cell("C3").Style.Font.FontColor = XLColor.Black;
        //        worksheet.Cell("D3").Style.Font.FontColor = XLColor.Black;
        //        worksheet.Cell("E3").Style.Font.FontColor = XLColor.Black;
        //        worksheet.Cell("F3").Style.Font.FontColor = XLColor.Black;
        //        worksheet.Cell("G3").Style.Font.FontColor = XLColor.Black;

        //        worksheet.ColumnWidth = 15;
        //        int mycount1 = 4;
        //        string mySql = "Select AccountCode,ParentAccountCode,AccountName,Prefix,CurrencyISOCode,ForeignCurrencyBalance,LocalCurrencyBalance,AddedDate from GLChartOfAccounts Where AccountStatus = 1 and ParentAccountCode=0";
        //        DataTable dtGLAccounts = DBUtils.GetDataTable(mySql);
        //        if (dtGLAccounts != null)
        //        {
        //            string pVal = "";
        //            foreach (DataRow MyRow in dtGLAccounts.Rows)
        //            {

        //                worksheet.Cell("A" + mycount1).Value ="P:"+ MyRow["AccountCode"].ToString();
        //                worksheet.Cell("B" + mycount1).Value = MyRow["AccountName"].ToString();
        //                worksheet.Cell("C" + mycount1).Value = MyRow["Prefix"].ToString();

        //                worksheet.Cell("D" + mycount1).Value = MyRow["CurrencyISOCode"].ToString();
        //                worksheet.Cell("E" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(Math.Abs(Common.toDecimal(MyRow["ForeignCurrencyBalance"])));
        //                worksheet.Cell("F" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(Math.Abs(Common.toDecimal(MyRow["LocalCurrencyBalance"])));
        //                worksheet.Cell("G" + mycount1).Value = Common.toDateTime(MyRow["AddedDate"]).ToString("yyyy-MM-dd");
        //                string abc = DBUtils.executeSqlGetSingle("Select AccountCode from GLChartOfAccounts Where AccountStatus = 1 and ParentAccountCode = '" + MyRow["AccountCode"].ToString() + "'");
        //                if (!string.IsNullOrEmpty(Common.toString(abc)))
        //                {
        //                    string sqlquery = "Select AccountCode, ParentAccountCode, AccountName,Prefix,CurrencyISOCode,ForeignCurrencyBalance,LocalCurrencyBalance,AddedDate from GLChartOfAccounts Where AccountStatus = 1 and ParentAccountCode = '" + MyRow["AccountCode"].ToString() + "'";
        //                    DataTable dtChildAccounts = DBUtils.GetDataTable(sqlquery);
        //                    if (dtChildAccounts != null)
        //                    {
        //                        foreach (DataRow MyRows in dtChildAccounts.Rows)
        //                        {
        //                            mycount1++;
        //                            worksheet.Cell("A" + mycount1).Value = MyRows["AccountCode"].ToString();
        //                            worksheet.Cell("B" + mycount1).Value = MyRows["AccountName"].ToString();
        //                            worksheet.Cell("C" + mycount1).Value = MyRows["Prefix"].ToString();

        //                            worksheet.Cell("D" + mycount1).Value = MyRows["CurrencyISOCode"].ToString();
        //                            worksheet.Cell("E" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(Math.Abs(Common.toDecimal(MyRows["ForeignCurrencyBalance"])));
        //                            worksheet.Cell("F" + mycount1).Value = DisplayUtils.GetSystemAmountFormat(Math.Abs(Common.toDecimal(MyRows["LocalCurrencyBalance"])));
        //                            worksheet.Cell("G" + mycount1).Value = Common.toDateTime(MyRows["AddedDate"]).ToString("yyyy-MM-dd");
        //                            worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.GreenYellow;
        //                            worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.GreenYellow;
        //                            worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.GreenYellow;
        //                            worksheet.Cell("D" + mycount1).Style.Fill.BackgroundColor = XLColor.GreenYellow;
        //                            worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.GreenYellow;
        //                            worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.GreenYellow;
        //                            worksheet.Cell("G" + mycount1).Style.Fill.BackgroundColor = XLColor.GreenYellow;

        //                            worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
        //                            worksheet.Cell("B" + mycount1).Style.Font.Bold = true;
        //                            worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
        //                            worksheet.Cell("D" + mycount1).Style.Font.Bold = true;
        //                            worksheet.Cell("E" + mycount1).Style.Font.Bold = true;
        //                            worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
        //                            worksheet.Cell("G" + mycount1).Style.Font.Bold = true;

        //                        }

        //                    }
        //                }
        //                worksheet.Cell("A" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
        //                worksheet.Cell("B" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
        //                worksheet.Cell("C" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
        //                worksheet.Cell("D" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
        //                worksheet.Cell("E" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
        //                worksheet.Cell("F" + mycount1).Style.Fill.BackgroundColor = XLColor.White;
        //                worksheet.Cell("G" + mycount1).Style.Fill.BackgroundColor = XLColor.White;

        //                worksheet.Cell("A" + mycount1).Style.Font.Bold = true;
        //                worksheet.Cell("B" + mycount1).Style.Font.Bold = true;
        //                worksheet.Cell("C" + mycount1).Style.Font.Bold = true;
        //                worksheet.Cell("D" + mycount1).Style.Font.Bold = true;
        //                worksheet.Cell("E" + mycount1).Style.Font.Bold = true;
        //                worksheet.Cell("F" + mycount1).Style.Font.Bold = true;
        //                worksheet.Cell("G" + mycount1).Style.Font.Bold = true;
        //                mycount1++;
        //            }
        //        }

        //        string fileName = Common.getRandomDigit(4) + "-" + "AccountBalances";
        //        workbook.SaveAs(Server.MapPath("~/Uploads/" + fileName + ".xlsx"));
        //        Response.Clear();
        //        Response.ClearHeaders();
        //        Response.ClearContent();
        //        Response.AddHeader("Content-Disposition", "attachment; filename= " + fileName + ".xlsx");
        //        Response.ContentType = "text/plain";
        //        Response.Flush();
        //        //Response.TransmitFile(Server.MapPath("~/Uploads/" + "/1/DeanRpt.xlsx"));
        //        Response.TransmitFile(Server.MapPath("~/Uploads/" + fileName + ".xlsx"));
        //        Response.End();
        //    }
        //    catch (Exception x)
        //    {
        //        Response.Write("<font color='Red'>" + x.Message + " </font>");

        //    }
        //}

        #endregion
    }
}
