using Application.DTOs;
using Dapper;
using Kendo.Mvc.UI;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace Accounting.Controllers
{
    public class UsersController : Controller
    {
        private AppDbContext _dbcontext = new AppDbContext();
        ////  private AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());
        //// private iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
        ApplicationUser Profile = ApplicationUser.GetUserProfile();

        private ApplicationSignInManager _signInManager;
        private ApplicationUserManager _userManager;

        public UsersController()
        {
        }

        public UsersController(ApplicationUserManager userManager, ApplicationSignInManager signInManager)
        {
            UserManager = userManager;
            SignInManager = signInManager;
        }

        public ApplicationSignInManager SignInManager
        {
            get
            {
                return _signInManager ?? HttpContext.GetOwinContext().Get<ApplicationSignInManager>();
            }
            private set
            {
                _signInManager = value;
            }
        }

        public ApplicationUserManager UserManager
        {
            get
            {
                return _userManager ?? HttpContext.GetOwinContext().GetUserManager<ApplicationUserManager>();
            }
            private set
            {
                _userManager = value;
            }
        }


        #region Administrator
        public ActionResult AdminList(UserModel User)
        {
            User.UserType = 99;
            return View("~/Views/Users/AdminList.cshtml", User);
        }
        public ActionResult AdminAdd()
        {
            var pUserModel = new RegisterViewModel();
            //Common.SyncUserMenus(0, "");
            pUserModel.menus = getMenuList("", 0);
            //ModelState.Clear();
            return View("~/Views/Users/AdminAdd.cshtml", pUserModel);
        }
        [HttpPost]
        public ActionResult AdminAdd(RegisterViewModel pUserModel)
        {
            if (ModelState.IsValid)
            {
                //if (pUserModel!=null)
                try
                {
                    pUserModel.UserType = 99;
                    pUserModel = AddUser(pUserModel);
                    if (pUserModel.isAdded)
                    {
                        pUserModel = saveNewUserMenuList(pUserModel);
                        return RedirectToAction("AdminList", "Users");
                    }
                    else
                    {
                        pUserModel.isAdded = false;
                        pUserModel.menus = getMenuList("", 0);
                        return View("~/Views/Users/AdminAdd.cshtml", pUserModel);
                    }
                }
                catch (Exception ex)
                {
                    //TempData["Failed"] = "Error occured: " + ex.Message.ToString();
                    pUserModel.Message = "Error occured: " + ex.Message.ToString();
                    pUserModel.isAdded = false;
                    pUserModel.menus = getMenuList("", 0);
                    return View("~/Views/Users/AdminAdd.cshtml", pUserModel);
                }
            }
            else
            {
                pUserModel.Message = Common.GetAlertMessage(1, "Invalid Data.");
                pUserModel.isAdded = false;
                pUserModel.menus = getMenuList("", 0);
                return View("~/Views/Users/AdminAdd.cshtml", pUserModel);
            }
        }

        public ActionResult AdminEdit(string Id)
        {
            //var users = db.Users.Where(x => x.UserID == Id).Take(1).FirstOrDefault();
            //UserModel UserM = new UserModel();
            //UserM.FirstName = users.FirstName;
            //UserM.LastName = users.LastName;
            //UserM.Email = users.Email;
            //UserM.FirstName = users.FirstName;
            //UserM.LastName = users.LastName;
            //UserM.DateOfBirth = users.DateOfBirth;
            //UserM.Address1 = users.Address1;
            //UserM.Address2 = users.Address2;
            //UserM.ZipCode = users.ZipCode;
            //UserM.CountryISOCode = Common.GetCountryName(users.CountryISOCode);
            //UserM.City = users.City;
            //UserM.LanguageISOCode = users.LanguageISOCode;
            //UserM.State = users.State;
            //UserM.PinCode = users.PinCode;
            //UserM.Fax = users.Fax;
            //UserM.CellPhone = users.CellPhone;
            //UserM.HomePhone = users.HomePhone;
            //UserM.OfficePhone = users.OfficePhone;
            //UserM.UserType = users.UserType;
            //UserM.UserID = users.UserID;
            //UserM.AddedDate = users.AddedDate;
            //UserM.Status = users.Status;

            UserModel UserM = new UserModel();
            string sql = "select * from Users where UserId='" + Id + "'";
            DataTable dt = DBUtils.GetDataTable(sql);
            if (dt != null && dt.Rows.Count > 0)
            {
                UserM.FirstName = Common.toString(dt.Rows[0]["FirstName"]);
                UserM.LastName = Common.toString(dt.Rows[0]["LastName"]);
                UserM.Email = Common.toString(dt.Rows[0]["Email"]);
                UserM.CountryISOCode = Common.toString(dt.Rows[0]["CountryISOCode"]);
                // UserM.LastName = Common.toString(dt.Rows[0]["LastName"]);
                UserM.DateOfBirth = Common.toDateTime(dt.Rows[0]["DateOfBirth"]);
                UserM.Address1 = Common.toString(dt.Rows[0]["Address1"]);
                UserM.Address2 = Common.toString(dt.Rows[0]["Address2"]);
                UserM.ZipCode = Common.toString(dt.Rows[0]["ZipCode"]);
                //UserM.CountryISOCode = Common.GetCountryName(Common.toString(dt.Rows[0]["CountryISOCode"]));
                UserM.City = Common.toString(dt.Rows[0]["City"]);
                UserM.LanguageISOCode = Common.toString(dt.Rows[0]["LanguageISOCode"]);
                UserM.State = Common.toString(dt.Rows[0]["State"]);
                UserM.PinCode = Common.toString(dt.Rows[0]["PinCode"]);
                UserM.Fax = Common.toString(dt.Rows[0]["Fax"]);
                UserM.CellPhone = Common.toString(dt.Rows[0]["CellPhone"]);
                UserM.HomePhone = Common.toString(dt.Rows[0]["HomePhone"]);
                UserM.OfficePhone = Common.toString(dt.Rows[0]["OfficePhone"]);
                UserM.UserType = Convert.ToByte(dt.Rows[0]["UserType"]);
                UserM.UserID = Common.toString(dt.Rows[0]["UserID"]);
                UserM.AddedDate = Common.toDateTime(dt.Rows[0]["AddedDate"]);
                UserM.Status = Common.toBool(dt.Rows[0]["Status"]);
            }
            return View("~/Views/Users/AdminEdit.cshtml", UserM);
        }
        [HttpPost]
        public ActionResult AdminEdit(UserModel pUserModel)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    pUserModel.UserType = 99;
                    pUserModel = UpdateUser(pUserModel);
                    if (pUserModel.isUpdated)
                    {
                        return RedirectToAction("AdminList", "Users");
                    }
                    else
                    {
                        pUserModel.isUpdated = false;
                        return View("~/Views/Users/AdminEdit.cshtml", pUserModel);
                    }
                }
                catch (Exception ex)
                {
                    //TempData["Failed"] = "Error occured: " + ex.Message.ToString();
                    pUserModel.Message = "Error occured: " + ex.Message.ToString();
                    pUserModel.isUpdated = false;
                    return View("~/Views/Users/AdminEdit.cshtml", pUserModel);
                }
            }
            else
            {
                pUserModel.Message = Common.GetAlertMessage(1, "Invalid Data.");
                pUserModel.isUpdated = false;
                return View("~/Views/Users/AdminEdit.cshtml", pUserModel);
            }
        }

        public ActionResult AdminDelete(string Id)
        {
            UserModel User = new UserModel();
            User.UserID = Id;
            return View("~/Views/Users/AdminDelete.cshtml", User);
        }
        [HttpPost]
        public ActionResult AdminDelete(UserModel User)
        {
            //// iRemitifyAccountsEntities db = new iRemitifyAccountsEntities();
            ////   var DelAdmin = db.Users.Where(x => x.UserID == User.UserID).Take(1).FirstOrDefault();
            string sql = "select UserID from Users where UserID='" + User.UserID + "'";
            string DelAdmin = Common.toString(DBUtils.executeSqlGetSingle(sql));
            if (DelAdmin != "")
            {
                DBUtils.ExecuteSQL("Delete Users where UserId='" + User.UserID + "'");
                ////  db.Users.Remove(DelAdmin);
                try
                {
                    TempData["Success"] = "User successfully deleted.";
                    return View("~/Views/Users/AdminDelete.cshtml", User);
                    //// db.SaveChanges();
                    // return RedirectToAction("AdminList", "Users");
                }
                catch (Exception ex)
                {
                    User.Message = "Error occured: " + ex.Message.ToString();
                    return View("~/Views/Users/AdminDelete.cshtml", User);
                }
            }
            else
            {
                User.Message = "Invalid User id.";
                return View("~/Views/Users/AdminDelete.cshtml", User);
            }
        }

        public ActionResult AdminMenu(string Id)
        {
            UserPermissionViewModel model = new UserPermissionViewModel();
            //  var objUser = db.Users.Where(x => x.UserID == Id).Take(1).FirstOrDefault();
            string sql = "select UserID from Users where UserID='" + Id + "'";
            string objUser = DBUtils.executeSqlGetSingle(sql);
            if (!string.IsNullOrEmpty(objUser))
            {
                model.UserID = objUser;
                Common.SyncUserMenus(0, Id);
                model.menus = getMenuList(Id, 0);
            }
            else
            {
                model.isValid = false;
                model.Message = "Invalid User";
            }
            return View("~/Views/Users/AdminMenu.cshtml", model);
        }
        [HttpPost]
        public ActionResult AdminMenu(UserPermissionViewModel model)
        {
            if (!string.IsNullOrEmpty(model.PinCode))
            {
                model.isValid = Common.ValidatePinCode(Profile.Id, model.PinCode);
                if (model.isValid)
                {
                    model = saveMenuList(model);
                }
                else
                {
                    model.Message = Common.GetAlertMessage(1, "Invalid pin code. Please provide valid pin code.");
                    model.menus = getMenuList(model.UserID, 0);

                }
            }
            else
            {
                model.Message = Common.GetAlertMessage(1, "Please provide valid pin code.");
                model.menus = getMenuList(model.UserID, 0);
            }
            return View("~/Views/Users/AdminMenu.cshtml", model);
        }
        #endregion

        #region Agent Managemet
        public ActionResult AgentList(AgentModel pUser)
        {
            pUser.UserType = 45;
            //GridDataItem item = e.Item as GridDataItem;
            //cpCustID = Convert.ToInt32(item.GetDataKeyValue("CustomerID"));
            DataTable myDatable = DBUtils.GetDataTable("Select CustomerFirstName,CustomerLastName,CustomerAddress1,CustomerAddress2,CustomerCity,CustomerState,CustomerZipCode,CustomerGender,CustomerCellPhone,CustomerOfficePhone,CustomerEmail,CustomerPlaceOfBirth,CustomerCompany,CustomerHomePhone From Customers Where CustomerID='" + pUser.CustomerID + "' ");
            if (myDatable.Rows.Count > 0)
            {
                //pnlCustomerInfo.Visible = true;
                DataRow myDR = myDatable.Rows[0];
                pUser.FirstName = myDR["CustomerFirstName"].ToString();//TypeID ?? 0);
                pUser.LastName = myDR["CustomerLastName"].ToString();
                pUser.Address1 = myDR["CustomerAddress1"].ToString();
                pUser.Address2 = myDR["CustomerAddress2"].ToString();
                pUser.City = myDR["CustomerCity"].ToString();
                pUser.State = myDR["CustomerState"].ToString();
                pUser.Gender = myDR["CustomerGender"].ToString() == "1" ? "Male" : "Female";
                pUser.CellPhone = myDR["CustomerCellPhone"].ToString();
                pUser.OfficePhone = myDR["CustomerOfficePhone"].ToString();
                pUser.Email = myDR["CustomerEmail"].ToString();
                pUser.DateOfBirth = Common.toDateTime(string.IsNullOrEmpty(myDR["CustomerPlaceOfBirth"].ToString()) ? "" : myDR["CustomerPlaceOfBirth"].ToString());
                pUser.ZipCode = myDR["CustomerZipCode"].ToString();
                pUser.Company = string.IsNullOrEmpty(myDR["CustomerCompany"].ToString()) ? "" : myDR["CustomerCompany"].ToString();
                pUser.HomePhone = string.IsNullOrEmpty(myDR["CustomerHomePhone"].ToString()) ? "" : myDR["CustomerHomePhone"].ToString();
                //rgCustomersList.Rebind();

                //RadGrid1.Rebind();
            }
            return View("~/Views/Users/Agents/AgentList.cshtml", pUser);
        }



        #endregion

        #region Generic User functions

        public ActionResult ListUsers([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM Users /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                          UserID, AddedDate, FirstName ,LastName,  Address1+' '+ Address2+' '+City+' '+ State +' '+ZipCode as FullAddress, CellPhone+' '+ HomePhone +' '+ OfficePhone as Phone,(select top 1 CountryName from ListCountries where ListCountries.CountryISONumericCode=Users.CountryISOCode ) as CountryISOCode,PhoneNumber ,Email from Users
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";

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
            else
            {
                builder.Where("UserType = 0");
            }

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("FirstName");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);
            var rows = _dbcontext.Query<UserModel>(selector.RawSql, selector.Parameters);
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }

        private RegisterViewModel AddUser(RegisterViewModel pUserModel)
        {
            pUserModel.isValid = true;
            var manager = HttpContext.GetOwinContext().GetUserManager<ApplicationUserManager>();
            var signInManager = HttpContext.GetOwinContext().Get<ApplicationSignInManager>();

            string strExist = DBUtils.executeSqlGetSingle("select Email from Users where email='" + pUserModel.Email.ToString() + "'");
            if (!string.IsNullOrEmpty(strExist))
            {
                pUserModel.isValid = false;
                pUserModel.Message = Common.GetAlertMessage(1, "Email already exist. Try another email.");
            }
            if (pUserModel.Password != pUserModel.ConfirmPassword)
            {
                pUserModel.isValid = false;
                pUserModel.Message = Common.GetAlertMessage(1, "Password and Confirm Password does not match.");
            }
            //else
            //{
            //    pUserModel.isValid = true;
            //}
            if (pUserModel.isValid)
            {
                var user = new ApplicationUser()
                {
                    UserName = Common.toString(pUserModel.Email),
                    Email = Common.toString(pUserModel.Email),
                    FirstName = Common.toString(pUserModel.FirstName),
                    LastName = Common.toString(pUserModel.LastName),
                    CountryISOCode = Common.toString(pUserModel.CountryISOCode),
                    DateOfBirth = pUserModel.DateOfBirth,
                    Address1 = Common.toString(pUserModel.Address1),
                    Address2 = Common.toString(pUserModel.Address2),
                    ZipCode = Common.toString(pUserModel.ZipCode),
                    City = Common.toString(pUserModel.City),
                    Fax = Common.toString(pUserModel.Fax),
                    State = Common.toString(pUserModel.State),
                    PinCode = Security.Encrypt(pUserModel.PinCode),
                    CellPhone = Common.toString(pUserModel.CellPhone),
                    HomePhone = Common.toString(pUserModel.HomePhone),
                    OfficePhone = Common.toString(pUserModel.OfficePhone),
                    UserType = pUserModel.UserType,
                    AddedDate = DateTime.Now,
                    Status = true
                };

                //user.ParentUserID = Profile.Id;

                //if (!string.IsNullOrEmpty(pUserModel.LanguageISOCode))
                //{
                //    user.LanguageISOCode = ddlUserLanguage.SelectedValue;
                //}
                if (!string.IsNullOrEmpty(pUserModel.CountryISOCode))
                {
                    user.CountryISOCode = pUserModel.CountryISOCode;
                }
                //pUserModel.State = user.State;
                //pUserModel.City = user.City;

                IdentityResult result = manager.Create(user, pUserModel.Password);

                if (result.Succeeded)
                {
                    pUserModel.isAdded = true;
                    pUserModel.UserID = user.Id;
                    if (!manager.IsInRole(user.Id, Common.getUserTypeName(Common.toString(pUserModel.UserType))))
                    {
                        manager.AddToRole(user.Id, Common.getUserTypeName(Common.toString(pUserModel.UserType)));

                    }
                }

                else
                {
                    pUserModel.Message = Common.GetAlertMessage(1, "Data missing. Please fill all required fields.");
                }
            }

            return pUserModel;
        }

        private UserModel UpdateUser(UserModel User)
        {
            if (ModelState.IsValid)
            {
                string strExist = DBUtils.executeSqlGetSingle("select Email from Users where email='" + User.Email.ToString() + "' and UserId not in ('" + User.UserID + "')");
                if (!string.IsNullOrEmpty(strExist))
                {
                    User.isValid = false;
                    User.Message = Common.GetAlertMessage(1, "Email already exist. Try another email.");
                }

                string UpdateSQL = "Update Users Set  Email= @Email,";
                UpdateSQL += " FirstName=@FirstName, LastName=@LastName, DateOfBirth=@DateOfBirth, Address1=@Address1,";
                UpdateSQL += " City =@City,CountryISOCode=@CountryISOCode, LanguageISOCode=@LanguageISOCode,State=@State,Fax=@Fax,CellPhone=@CellPhone, HomePhone =@HomePhone, OfficePhone=@OfficePhone, ZipCode=@ZipCode ";
                UpdateSQL += " Where UserID =@UserID";

                using (SqlConnection con = Common.getConnection())
                {
                    SqlCommand updateCommand = con.CreateCommand();
                    try
                    {
                        //updateCommand.Parameters.AddWithValue("@UserName", User.UserName);
                        updateCommand.Parameters.AddWithValue("@Email", Common.toString(User.Email));
                        updateCommand.Parameters.AddWithValue("@FirstName", Common.toString(User.FirstName));
                        updateCommand.Parameters.AddWithValue("@LastName", Common.toString(User.LastName));
                        updateCommand.Parameters.AddWithValue("@DateOfBirth", Common.toDateTime(User.DateOfBirth));
                        updateCommand.Parameters.AddWithValue("@Address1", Common.toString(User.Address1));
                        updateCommand.Parameters.AddWithValue("@Fax", Common.toString(User.Fax));
                        updateCommand.Parameters.AddWithValue("@City", Common.toString(User.City));
                        updateCommand.Parameters.AddWithValue("@CountryISOCode", Common.toString(User.CountryISOCode));
                        updateCommand.Parameters.AddWithValue("@State", Common.toString(User.State));
                        updateCommand.Parameters.AddWithValue("@CellPhone", Common.toString(User.CellPhone));
                        updateCommand.Parameters.AddWithValue("@HomePhone", Common.toString(User.HomePhone));

                        updateCommand.Parameters.AddWithValue("@OfficePhone", Common.toString(User.OfficePhone));
                        updateCommand.Parameters.AddWithValue("@UserID", Common.toString(User.UserID));
                        updateCommand.Parameters.AddWithValue("@LanguageISOCode", Common.toString(User.LanguageISOCode));
                        updateCommand.Parameters.AddWithValue("@ZipCode", Common.toString(User.ZipCode));
                        //updateCommand.Parameters.AddWithValue("@AddedDate", User.DateAdded);
                        updateCommand.CommandText = UpdateSQL;
                        int RowsAffected = updateCommand.ExecuteNonQuery();
                        if (RowsAffected == 1)
                        {
                            User.isUpdated = true;
                            User.Message = Common.GetAlertMessage(1, "User profile successfully updated.");
                        }
                        else
                        {
                            User.Message = Common.GetAlertMessage(0, "Data missing. Please fill all required fields.");
                        }

                    }
                    catch (Exception ex)
                    {
                        User.Message = Common.GetAlertMessage(1, "Data missing. Please fill all required fields.");
                    }
                    finally
                    {
                        con.Close();
                    }
                }
            }
            else
            {
                User.Message = Common.GetAlertMessage(1, "Email already exist .");
            }
            return User;
        }
        public ActionResult UpdatePassword(string Id)
        {
            // var objUser = db.Users.Where(x => x.UserID == Id).Take(1).FirstOrDefault();

            ResetPasswordDto userModel = new ResetPasswordDto();
            string sql = "select UserID,FirstName,LastName,Email,AKey from Users where UserID='" + Id + "'";
            DataTable objUser = DBUtils.GetDataTable(sql);
            if (objUser != null && objUser.Rows.Count > 0)
            {
                userModel.UserID = objUser.Rows[0]["UserID"].ToString();
                userModel.FullName = objUser.Rows[0]["FirstName"].ToString() + " " + objUser.Rows[0]["LastName"].ToString();
                if (!string.IsNullOrEmpty(objUser.Rows[0]["AKey"].ToString()))
                {
                    userModel.AKey = Security.Decrypt(objUser.Rows[0]["AKey"].ToString());
                }
                userModel.Email = objUser.Rows[0]["Email"].ToString();
            }
            return View("~/Views/Users/shared/_UpdatePassword.cshtml", userModel);
        }
        public ActionResult UserSecuritySettings(string Id)
        {
            // var objUser = db.Users.Where(x => x.UserID == Id).Take(1).FirstOrDefault();

            UserModel userModel = new UserModel();
            string sql = "select UserID,FirstName,LastName,Email,AKey from Users where UserID='" + Id + "'";
            DataTable objUser = DBUtils.GetDataTable(sql);
            if (objUser != null && objUser.Rows.Count > 0)
            {
                userModel.UserID = objUser.Rows[0]["UserID"].ToString();
            }
            return View("~/Views/Users/UserSecuritySettings.cshtml", userModel);
        }



        [HttpPost]

        public async Task<ActionResult> UpdatePassword(ResetPasswordDto pUserModel)
        {
            //string strExist = DBUtils.executeSqlGetSingle("select Email from Users where email='" + pUserModel.Email.ToString() + "'");
            //if (!string.IsNullOrEmpty(strExist))
            //{
            //    pUserModel.isValid = false;
            //    pUserModel.Message = Common.GetAlertMessage(1, "Email already exist. Try another email.");
            //}
            if (pUserModel.Password != pUserModel.ConfirmPassword && pUserModel.Password != "")
            {
                pUserModel.isValid = false;
                pUserModel.Message = Common.GetAlertMessage(1, "Password and confirm password does not match or password is empty.");
            }
            else
            {
                pUserModel.isValid = true;
            }
            if (pUserModel.isValid)
            {
                if (!string.IsNullOrEmpty(pUserModel.Password))
                {
                    var user = await UserManager.FindByEmailAsync(pUserModel.Email);
                    if (user != null)
                    {
                        var result = await UserManager.RemovePasswordAsync(pUserModel.UserID.ToString());
                        if (result.Succeeded)
                        {
                            result = await UserManager.AddPasswordAsync(pUserModel.UserID, pUserModel.Password);
                            if (result.Succeeded)
                            {
                                string NewAkey = Security.Encrypt(pUserModel.Password);
                                if (!string.IsNullOrEmpty(pUserModel.UserID))
                                {
                                    bool IsUpdated = false;
                                    string sqlstring = "update users set Akey = '" + NewAkey + "' where userid = '" + pUserModel.UserID + "'";
                                    IsUpdated = DBUtils.ExecuteSQL(sqlstring);
                                    if (IsUpdated)
                                    {
                                        if (pUserModel.IsSendEmail)
                                        {
                                            string Subject = "Change Password";
                                            string Body = " Dear " + pUserModel.FullName + ", <br>";
                                            Body += " Your password has been updated. Your new password is " + pUserModel.Password + "<br>";
                                            Common.SendEmail(pUserModel.Email, Profile.Email, Subject, Body);
                                            pUserModel.Message = Common.GetAlertMessage(0, "Your change password email sent successfully");
                                            return View("~/Views/Users/shared/_UpdatePassword.cshtml", pUserModel);
                                        }
                                        pUserModel.Message = Common.GetAlertMessage(0, "Your change password changed successfully");
                                        return View("~/Views/Users/shared/_UpdatePassword.cshtml", pUserModel);
                                    }

                                }
                            }
                            else
                            {
                                pUserModel.Message = Common.GetAlertMessage(1, "Password does not updated. Please try again.");
                                return View("~/Views/Users/shared/_UpdatePassword.cshtml", pUserModel);
                            }
                        }
                        else
                        {
                            pUserModel.Message = Common.GetAlertMessage(1, "Data missing. Please check all fields and try again.");
                            return View("~/Views/Users/shared/_UpdatePassword.cshtml", pUserModel);
                        }
                    }
                    else
                    {
                        pUserModel.Message = Common.GetAlertMessage(1, "Data missing. Please check all fields and try again.");
                        return View("~/Views/Users/shared/_UpdatePassword.cshtml", pUserModel);
                    }
                }
                else
                {
                    pUserModel.Message = Common.GetAlertMessage(1, "Please provide new password.");
                    return View("~/Views/Users/shared/_UpdatePassword.cshtml", pUserModel);
                }
            }
            return View("~/Views/Users/shared/_UpdatePassword.cshtml", pUserModel);

        }


        public ActionResult UpdatePincode(string Id)
        {
            // var users = db.Users.Where(x => x.UserID == Id).Take(1).FirstOrDefault();
            UserModel UserM = new UserModel();
            UserM.UserID = Id;
            return View("~/Views/Users/shared/_UpdatePincode.cshtml", UserM);
        }
        [HttpPost]
        public ActionResult UpdatePincode(UserModel User)
        {
            if (!string.IsNullOrEmpty(User.PinCode) && !string.IsNullOrEmpty(User.OldPinCode))
            {
                if (User.UserID.Trim() != "" && User.OldPinCode != "")
                {
                    if (!Common.ValidatePinCode(Profile.Id, User.OldPinCode.Trim()))
                    {
                        User.Message = Common.GetAlertMessage(1, "Invalid Pin code entered");
                    }
                    else
                    {
                        string CustomerFullName = "";
                        string SqlUpdateUser = "Update [Users] set PinCode=@PinCode,PinCodeExpiryDate=@PinCodeExpiryDate where UserID=@UserID";
                        string SqlGetTemplate = "Select TemplateCode,TemplateSubject,TemplateBody,TemplateName from DefaultTemplates where TemplateCode='UPCBU'";
                        DataTable GetTemplate = DBUtils.GetDataTable(SqlGetTemplate);

                        string SqlGetUser = "Select FirstName,LastName,Email from [Users] where UserID='" + User.UserID + "'";
                        DataTable GetUser = DBUtils.GetDataTable(SqlGetUser);
                        DataRow UserInfo = null;
                        if (GetUser != null && GetUser.Rows.Count >= 0)
                        {
                            UserInfo = GetUser.Rows[0];
                            CustomerFullName = UserInfo["FirstName"].ToString() + " " + UserInfo["LastName"].ToString();
                        }

                        using (SqlConnection con = Common.getConnection())
                        {
                            try
                            {
                                SqlCommand UpdateUserCommand = con.CreateCommand();

                                string GetRandom = Security.Encrypt(User.PinCode); //Utils.getRandomDigit(4);
                                UpdateUserCommand.Parameters.AddWithValue("@PinCode", GetRandom);
                                UpdateUserCommand.Parameters.AddWithValue("@PinCodeExpiryDate", DateTime.UtcNow.AddYears(2));
                                UpdateUserCommand.Parameters.AddWithValue("@UserID", User.UserID);

                                UpdateUserCommand.CommandText = SqlUpdateUser;
                                UpdateUserCommand.ExecuteNonQuery();
                                UpdateUserCommand.Parameters.Clear();
                                if (GetTemplate != null && GetTemplate.Rows.Count >= 0)
                                {
                                    DataRow Template = GetTemplate.Rows[0];
                                    if (!string.IsNullOrEmpty(Template["TemplateSubject"].ToString()) && !string.IsNullOrEmpty(Template["TemplateBody"].ToString()))
                                    {
                                        string strTemplateBody = Template["TemplateBody"].ToString();
                                        strTemplateBody = strTemplateBody.Replace("##CustomerName##", CustomerFullName);
                                        strTemplateBody = strTemplateBody.Replace("##PinCode##", CustomerFullName);
                                        string strTemplateSubject = Template["TemplateSubject"].ToString();
                                        Common.SendEmail(UserInfo["Email"].ToString(), SysPrefs.AutoEmail, strTemplateSubject, strTemplateBody);
                                    }

                                    User.Message = Common.GetAlertMessage(0, "User pin code updated successfully. <br/> <a href = 'javascript: RefreshParent();'>Click here to close</a>");
                                }
                            }
                            catch (Exception ex)
                            {
                                User.Message = Common.GetAlertMessage(1, "Failed to update User Pin code.Please try again" + ex.Message.ToString());
                            }
                            finally
                            {
                                con.Close();
                            }
                        }
                    }
                }
                else
                {
                    User.Message = Common.GetAlertMessage(1, "Invalid user id or invalid pin code entered <br/> <a href = 'javascript: RefreshParent();'>Click here to close</a>");
                }
            }
            else
            {
                User.Message = Common.GetAlertMessage(1, "Please provide new pin code and old pin code. <br/> <a href = 'javascript: RefreshParent();'>Click here to close</a>");
            }
            return View("~/Views/Users/shared/_UpdatePincode.cshtml", User);
        }

        public ActionResult UpdatePhoto(string Id)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            UserModel UserM = new UserModel();
            // var ObjUsers = db.Users.Where(x => x.UserID == Id).Take(1).FirstOrDefault();
            string sql = "select UserID,Photo from Users where UserID='" + Id + "'";
            DataTable objUser = DBUtils.GetDataTable(sql);
            if (objUser != null && objUser.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(objUser.Rows[0]["Photo"].ToString()))
                {
                    UserM.UserID = Id;
                    UserM.ImageUrl = "~/uploads/" + SysPrefs.SubmissionFolder + "/profiles/" + objUser.Rows[0]["UserID"] + "/" + objUser.Rows[0]["Photo"];
                }
                else
                {
                    UserM.UserID = Id;

                    UserM.ImageUrl = "~/assets/img/user.jpg";
                }
            }
            else
            {
                UserM.Message = "Invalid User";
            }
            return View("~/Views/Users/shared/_UpdatePhoto.cshtml", UserM);
        }
        [HttpPost]
        public ActionResult UpdatePhoto(UserModel User, HttpPostedFileBase file)
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            if (file != null && file.ContentLength > 0)
            {
                string SqlUpDateUser = "Update [Users] set Photo=@Photo where UserID=@UserID";
                string fileName = "";
                if (file.ContentLength > 0)
                {
                    string savePath = "";
                    fileName = file.FileName;
                    fileName = Common.getRandomString(5) + fileName.Replace(@"[^\w\.@-]`~!@#$%^&*()=+;", "");
                    ////savePath = Request.PhysicalApplicationPath + Settings.UserProfileFolder.ToString() + Profile.AiuUserID.ToString() + "/Assignments";
                    savePath = Server.MapPath("~/uploads/" + SysPrefs.SubmissionFolder + "/profiles/" + User.UserID + "");
                    if (!Directory.Exists(savePath))
                    {
                        Directory.CreateDirectory(savePath);
                    }
                    savePath += "/" + fileName;
                    file.SaveAs(savePath);

                    using (SqlConnection con = Common.getConnection())
                    {
                        try
                        {
                            SqlCommand Updateuser = con.CreateCommand();

                            Updateuser.Parameters.AddWithValue("@Photo", fileName);
                            Updateuser.Parameters.AddWithValue("@UserID", User.UserID);
                            Updateuser.CommandText = SqlUpDateUser;
                            Updateuser.ExecuteNonQuery();
                            Updateuser.Parameters.Clear();
                            User.Message = Common.GetAlertMessage(0, "Profile Image Updated Successfully.<br/> <a href = 'javascript: RefreshParent();'>Click here to close</a>");
                            User.ImageUrl = "/uploads/" + SysPrefs.SubmissionFolder + "/profiles/" + User.UserID + "/" + fileName;
                        }
                        catch
                        {
                            User.Message = Common.GetAlertMessage(1, "failed to update Profile Image. Please try again.");
                        }
                        finally
                        {
                            con.Close();
                        }
                    }
                }
                else
                {
                    User.Message = Common.GetAlertMessage(1, "failed to update Profile Image. Please try again.");
                }
            }
            else
            {
                User.Message = Common.GetAlertMessage(1, "Please select a file to upload.");
                User.ImageUrl = "~/assets/img/user.jpg";
            }
            return View("~/Views/Users/shared/_UpdatePhoto.cshtml", User);
        }


        #endregion

        #region Set User IPAddress 
        public ActionResult SetUsersIPAddress(string Id)
        {
            UsersIPAdressDto userModel = new UsersIPAdressDto();
            string sql = "select UserID,FirstName,LastName,Email,AKey,VerifyIPAddress from Users where UserID='" + Id + "'";
            DataTable objUser = DBUtils.GetDataTable(sql);
            if (objUser != null && objUser.Rows.Count > 0)
            {
                userModel.UserID = objUser.Rows[0]["UserID"].ToString();
                userModel.VerifyIPAddress = Common.toBool(objUser.Rows[0]["VerifyIPAddress"]);
                userModel.Status = true;
            }
            return View("~/Views/Users/SetUsersIPAddress.cshtml", userModel);
        }
        public ActionResult ListIPAddress([DataSourceRequest] DataSourceRequest request)
        {
            const string countQuery = @"SELECT COUNT(1) FROM UsersIPAdress /**where**/";
            const string selectQuery = @"SELECT  *
                           FROM    ( SELECT    ROW_NUMBER() OVER ( /**orderby**/ ) AS RowNum, 
                          UsersIPAdressID,UserID,IPAddress, DateAdded,Status from UsersIPAdress
                                     /**where**/  
                                   ) AS RowConstrainedResult
                           WHERE   RowNum >= (@PageIndex * @PageSize + 1 )
                               AND RowNum <= (@PageIndex + 1) * @PageSize
                           ORDER BY RowNum";

            SqlBuilder builder = new SqlBuilder();
            var count = builder.AddTemplate(countQuery);

            int CurrentPage = 0; int CurrentPageSize = 25;
            if (request.PageSize > 0)
            {
                //   CurrentPageSize = request.PageSize;
            }
            if (request.PageSize > 0 && request.Page > 0)
            {
                // CurrentPage = (request.Page - 1);
            }
            var selector = builder.AddTemplate(selectQuery, new { PageIndex = CurrentPage, PageSize = CurrentPageSize });

            if (request.Filters != null && request.Filters.Any())
            {
                builder = Common.ApplyFilters(builder, request.Filters);
            }
            else
            {
                //builder.Where("UserID > 0");
            }

            if (request.Sorts != null && request.Sorts.Any())
            {
                builder = Common.ApplySorting(builder, request.Sorts);
            }
            else
            {
                builder.OrderBy("UsersIPAdressID");
            }

            var totalCount = _dbcontext.QueryFirst<int>(count.RawSql, count.Parameters);
            var rows = _dbcontext.Query<UsersIPAdressDto>(selector.RawSql, selector.Parameters);
            var result = new DataSourceResult()
            {
                Data = rows,
                Total = totalCount
            };
            return Json(result);
        }
        public ActionResult AddUserIPAddress(string Id)
        {
            UsersIPAdressDto pModel = new UsersIPAdressDto();
            if (Id != null)
            {
                string sql = "select * from UsersIPAdress where UsersIPAdressID='" + Id + "'";
                DataTable objUser = DBUtils.GetDataTable(sql);
                if (objUser != null && objUser.Rows.Count > 0)
                {
                    pModel.UsersIPAdressID = Common.toString(objUser.Rows[0]["UsersIPAdressID"]);
                    pModel.UserID = objUser.Rows[0]["UserID"].ToString();
                    pModel.IPAddress = objUser.Rows[0]["IPAddress"].ToString();
                    pModel.Status = Common.toBool(objUser.Rows[0]["Status"]);
                    pModel.VerifyIPAddress = Common.toBool(DBUtils.executeSqlGetSingle("select VerifyIPAddress from Users  where UserId='" + pModel.UserID + "'"));
                }

            }
            else
            {
                pModel.Message = Common.GetAlertMessage(1, "Data missing. Invalid user information.");
                return View("~/Views/Users/", pModel);
            }
            return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
        }
        [HttpPost]
        public ActionResult AddUserIPAddress(UsersIPAdressDto pModel)
        {
            bool IsIP = false;
            pModel.VerifyIPAddress = Common.toBool(DBUtils.executeSqlGetSingle("select VerifyIPAddress from Users  where UserId='" + pModel.UserID + "'"));
            if (ModelState.IsValid)
            {
                if (pModel.UserID != null)
                {
                    pModel.VerifyIPAddress = Common.toBool(DBUtils.executeSqlGetSingle("select VerifyIPAddress from Users  where UserId='" + pModel.UserID + "'"));
                    if (!string.IsNullOrEmpty(pModel.IPAddress))
                    {

                        Regex regexIP = new Regex(@"^([01]?[0-9]?[0-9]|2[0-4][0-9]|25[0-5])\.([01]?[0-9]?[0-9]|2[0-4][0-9]|25[0-5])\.([01]?[0-9]?[0-9]|2[0-4][0-9]|25[0-5])\.([01]?[0-9]?[0-9]|2[0-4][0-9]|25[0-5])$");

                        if (regexIP.Match(pModel.IPAddress).Success)
                        {
                            IsIP = true;
                        }
                        //if (!string.IsNullOrEmpty(pModel.UsersIPAdressID))
                        //{
                        //    string sqlUpdate = "update UsersIPAdress set IPAddress='" + pModel.IPAddress + "',status='1' where UsersIPAdressID=" + pModel.UsersIPAdressID + "";
                        //    DBUtils.ExecuteSQL(sqlUpdate);
                        //    pModel.Message = Common.GetAlertMessage(0, "successfully updated..");
                        //    return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
                        //}
                        //else
                        //{
                        if (IsIP)
                        {
                            string sqlUpdate = "insert into UsersIPAdress (UserID,IPAddress,Status,DateAdded)";
                            sqlUpdate += " Values('" + pModel.UserID + "','" + pModel.IPAddress + "','1','" + DateTime.Now.ToString("yyyy-MM-dd 00:00:00") + "')";
                            DBUtils.ExecuteSQL(sqlUpdate);
                            pModel.Message = Common.GetAlertMessage(0, "successfully added..");
                            return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
                        }
                        else
                        {
                            pModel.Message = Common.GetAlertMessage(1, "Invalid IP Address.");
                            return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
                        }
                        //}
                    }
                    else
                    {
                        pModel.Message = Common.GetAlertMessage(1, "Data missing. Invalid user information.");
                        return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
                    }
                }
                else
                {
                    pModel.Message = Common.GetAlertMessage(1, "Data missing. Invalid user information.");
                    return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
                }
            }
            return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
        }
        public ActionResult DeleteUserIPAddress(string Id)
        {
            UsersIPAdressDto pModel = new UsersIPAdressDto();
            if (!string.IsNullOrEmpty(Id))
            {
                pModel.UserID = DBUtils.executeSqlGetSingle("select UserID from UsersIPAdress  where UsersIPAdressID=" + Id + "");
                pModel.VerifyIPAddress = Common.toBool(DBUtils.executeSqlGetSingle("select VerifyIPAddress from Users  where UserId='" + pModel.UserID + "'"));
                string sqlUpdate = "delete UsersIPAdress where UsersIPAdressID=" + Id + "";
                DBUtils.ExecuteSQL(sqlUpdate);
                pModel.Message = Common.GetAlertMessage(0, "successfully deleted.");
                return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
            }
            else
            {
                pModel.Message = Common.GetAlertMessage(1, "Data missing. Invalid user information.");
                return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
            }
            return View("~/Views/Users/SetUsersIPAddress.cshtml", pModel);
        }
        public void VerifyIPAddress(int VerifyIPAddress, string userid)
        {
            if (!string.IsNullOrEmpty(userid))
            {
                string sqlUpdate = "update Users set VerifyIPAddress='" + VerifyIPAddress + "' where UserId='" + userid + "'";
                DBUtils.ExecuteSQL(sqlUpdate);
            }
        }

        #endregion

        #region Helper functions
        public static List<UserMenuDto> getMenuList(string pUserId, int ParentMenuId)
        {
            ////AppDbContext _dbcontext = new AppDbContext(BaseModel.getConnString());
            AppDbContext _dbcontext = new AppDbContext();
            string selectQuery = @"SELECT * FROM( SELECT UsersMenu.UserID, DefaultMenus.MenuID, DefaultMenus.Title, UsersMenu.Status as isAssigned, UsersMenu.ActionControls FROM DefaultMenus Left OUTER JOIN UsersMenu ON DefaultMenus.MenuID = UsersMenu.MenuID /**where**/ ) AS RowConstrainedResult";
            SqlBuilder builder = new SqlBuilder();
            var selector = builder.AddTemplate(selectQuery, new { PageIndex = 0, PageSize = 300 });

            builder.Where("DefaultMenus.UserType = 99");
            builder.Where("DefaultMenus.Status = 1");
            builder.Where("DefaultMenus.VisibleInMenu = 1");
            builder.Where("DefaultMenus.ParentMenuId = " + ParentMenuId);
            builder.Where("UsersMenu.userId = '" + pUserId + "'");
            builder.OrderBy("DefaultMenus.SortOrder");
            var rows = _dbcontext.Query<UserMenuDto>(selector.RawSql, selector.Parameters);
            if (rows.Count() == 0)
            {
                selectQuery = @"SELECT * FROM(SELECT '' as  UserID, DefaultMenus.MenuID, DefaultMenus.Title, DefaultMenus.Status, '' as ActionControls FROM DefaultMenus /**where**/ ) AS RowConstrainedResult";
                builder = new SqlBuilder();
                selector = builder.AddTemplate(selectQuery, new { PageIndex = 0, PageSize = 300 });
                builder.Where("DefaultMenus.UserType = 99");
                builder.Where("DefaultMenus.Status = 1");
                builder.Where("DefaultMenus.VisibleInMenu = 1");
                builder.Where("DefaultMenus.ParentMenuId = " + ParentMenuId);
                builder.OrderBy("DefaultMenus.SortOrder");
                rows = _dbcontext.Query<UserMenuDto>(selector.RawSql, selector.Parameters);
            }
            foreach (UserMenuDto model in rows)
            {
                model.SubMenu = getMenuList(pUserId, model.MenuID);
            }
            return rows.ToList();
        }

        public UserPermissionViewModel saveMenuList(UserPermissionViewModel model)
        {
            using (SqlConnection con = Common.getConnection())
            {
                try
                {
                    for (var i = 0; i < model.menus.Count(); i++)
                    {
                        string sqlParent = "Update UsersMenu set Status = '" + model.menus[i].isAssigned + "' where MenuID = '" + model.menus[i].MenuID + "' and UserID = '" + model.UserID + "'";
                        SqlCommand updateParent = con.CreateCommand();
                        updateParent.CommandText = sqlParent;
                        updateParent.ExecuteNonQuery();
                        updateParent.Parameters.Clear();
                        if (model.menus[i].subMenu != null)
                        {
                            if (model.menus[i].subMenu.Count > 0)
                            {
                                for (var j = 0; j < model.menus[i].subMenu.Count(); j++)
                                {
                                    string sqlChild = "Update UsersMenu set Status = '" + model.menus[i].subMenu[j].isAssigned + "' where MenuID = '" + model.menus[i].subMenu[j].MenuID + "' and UserID = '" + model.UserID + "'";
                                    SqlCommand updateChild = con.CreateCommand();
                                    updateChild.CommandText = sqlChild;
                                    updateChild.ExecuteNonQuery();
                                    updateChild.Parameters.Clear();
                                    if (model.menus[i].subMenu[j].subMenu != null)
                                    {
                                        if (model.menus[i].subMenu[j].subMenu.Count > 0)
                                        {
                                            for (var k = 0; k < model.menus[i].subMenu[j].subMenu.Count(); k++)
                                            {
                                                string sqlSubChild = "Update UsersMenu set Status = '" + model.menus[i].subMenu[j].subMenu[k].isAssigned + "' where MenuID = '" + model.menus[i].subMenu[j].subMenu[k].MenuID + "' and UserID = '" + model.UserID + "'";
                                                SqlCommand updateSubChild = con.CreateCommand();
                                                updateSubChild.CommandText = sqlSubChild;
                                                updateSubChild.ExecuteNonQuery();
                                                updateSubChild.Parameters.Clear();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    model.Message = Common.GetAlertMessage(0, "User permissions updated successfully. <br/> <a href = 'javascript: RefreshParent();'>Click here to close</a>");
                }
                catch (Exception ex)
                {
                    model.Message = Common.GetAlertMessage(1, "Failed to update User permissions.Please try again" + ex.Message.ToString());
                }
                finally
                {
                    con.Close();
                }
            }
            model.menus = getMenuList(model.UserID, 0);
            return model;
        }

        public static RegisterViewModel saveNewUserMenuList(RegisterViewModel model)
        {
            using (SqlConnection con = Common.getConnection())
            {
                try
                {
                    if (model.menus != null)
                    {
                        for (var i = 0; i < model.menus.Count(); i++)
                        {
                            string sqlParent = "Insert into UsersMenu (UserID,MenuID, Status) values('" + model.UserID + "','" + model.menus[i].MenuID + "','" + model.menus[i].isAssigned + "')";
                            SqlCommand updateParent = con.CreateCommand();
                            updateParent.CommandText = sqlParent;
                            updateParent.ExecuteNonQuery();
                            updateParent.Parameters.Clear();
                            if (model.menus[i].subMenu != null)
                            {
                                if (model.menus[i].subMenu.Count > 0)
                                {
                                    for (var j = 0; j < model.menus[i].subMenu.Count(); j++)
                                    {
                                        string sqlChild = "Insert into UsersMenu (UserID, MenuID, Status) values('" + model.UserID + "', '" + model.menus[i].subMenu[j].MenuID + "', '" + model.menus[i].subMenu[j].isAssigned + "')";
                                        SqlCommand updateChild = con.CreateCommand();
                                        updateChild.CommandText = sqlChild;
                                        updateChild.ExecuteNonQuery();
                                        updateChild.Parameters.Clear();
                                        if (model.menus[i].subMenu[j].subMenu != null)
                                        {
                                            if (model.menus[i].subMenu[j].subMenu.Count > 0)
                                            {
                                                for (var k = 0; k < model.menus[i].subMenu[j].subMenu.Count(); k++)
                                                {
                                                    string sqlSubChild = "Insert into UsersMenu (UserID, MenuID, Status) values('" + model.UserID + "', '" + model.menus[i].subMenu[j].subMenu[k].MenuID + "', '" + model.menus[i].subMenu[j].subMenu[k].isAssigned + "')";
                                                    SqlCommand updateSubChild = con.CreateCommand();
                                                    updateSubChild.CommandText = sqlSubChild;
                                                    updateSubChild.ExecuteNonQuery();
                                                    updateSubChild.Parameters.Clear();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    model.Message = Common.GetAlertMessage(0, "User permissions updated successfully. <br/> <a href = 'javascript: RefreshParent();'>Click here to close</a>");
                }
                catch (Exception ex)
                {
                    model.Message = Common.GetAlertMessage(1, "Failed to update User permissions.Please try again" + ex.Message.ToString());
                }
                finally
                {
                    con.Close();
                }
            }
            model.menus = getMenuList(model.UserID, 0);

            return model;
        }

        #endregion
    }
}