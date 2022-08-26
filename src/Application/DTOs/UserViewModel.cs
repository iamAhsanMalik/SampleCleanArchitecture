using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using static DHAAccounts.Models.Enum;

namespace Application.DTOs
{
    public class UserViewModel
    {
        public List<UserModel> userViewModel = new List<UserModel>();
    }
    public class AgentModel
    {
        public string CustomerID { get; set; }
        public string Company { get; set; }
        public string ImageUrl { get; set; }
        public string UserID { get; set; }
        [Required(ErrorMessage = "Email is required")]
        public string Email { get; set; }
        public string Password { get; set; }
        public string UserName { get; set; }
        [Required(ErrorMessage = "First name is required")]
        public string FirstName { get; set; }
        [Required(ErrorMessage = "Last name is required")]
        public string LastName { get; set; }
        [Required(ErrorMessage = "Date of birth is required")]
        public DateTime? DateOfBirth { get; set; }
        public string Gender { get; set; }
        public string FullAddress { get; set; }
        [Required(ErrorMessage = "Address is required")]
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        [Required(ErrorMessage = "City is required")]
        public string City { get; set; }
        [Required(ErrorMessage = "State is required")]
        public string State { get; set; }
        [Required(ErrorMessage = "Zip code is required")]
        public string ZipCode { get; set; }
        [Required(ErrorMessage = "Country iso code is required")]
        public string CountryISOCode { get; set; }
        public string Phone { get; set; }
        [Required(ErrorMessage = "Cell phone is required")]
        public string CellPhone { get; set; }
        public string HomePhone { get; set; }
        public string OfficePhone { get; set; }
        [Required(ErrorMessage = "Fax is required")]
        public string Fax { get; set; }
        public byte? UserType { get; set; }
        [Required(ErrorMessage = "Language iso code is required")]
        public string LanguageISOCode { get; set; }
        public bool? Status { get; set; }
        public string Photo { get; set; }
        public bool VerifyIPAddress { get; set; }
        public string AddedBy { get; set; }
        public DateTime? AddedDate { get; set; }
        public string PinCode { get; set; }
        public string OldPinCode { get; set; }
        public DateTime PinCodeExpiryDate { get; set; }
        public bool isValid { get; set; }
        public string Message { get; set; }
        public bool isUpdated { get; set; }

    }
    public class UserModel
    {
        public string ImageUrl { get; set; }
        public string UserID { get; set; }
        [Required(ErrorMessage = "Email is required")]
        public string Email { get; set; }
        public string Password { get; set; }

        public string UserName { get; set; }
        [Required(ErrorMessage = "First name is required")]
        public string FirstName { get; set; }
        [Required(ErrorMessage = "Last name is required")]
        public string LastName { get; set; }
        //[Required(ErrorMessage = "Date of birth is required")]
        public DateTime? DateOfBirth { get; set; }
        public string Gender { get; set; }
        public string FullAddress { get; set; }
        //[Required(ErrorMessage = "Address is required")]
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        //[Required(ErrorMessage = "City is required")]
        public string City { get; set; }
        //[Required(ErrorMessage = "State is required")]
        public string State { get; set; }
        //[Required(ErrorMessage = "Zip code is required")]
        public string ZipCode { get; set; }
        //[Required(ErrorMessage = "Country iso code is required")]
        public string CountryISOCode { get; set; }
        public string Phone { get; set; }
        //[Required(ErrorMessage = "Cell phone is required")]
        public string CellPhone { get; set; }

        public string HomePhone { get; set; }
        public string OfficePhone { get; set; }
        //[Required(ErrorMessage = "Fax is required")]
        public string Fax { get; set; }
        public byte? UserType { get; set; }
        //[Required(ErrorMessage = "Language iso code is required")]
        public string LanguageISOCode { get; set; }
        public bool? Status { get; set; }
        public string Photo { get; set; }
        public bool VerifyIPAddress { get; set; }
        public string AddedBy { get; set; }
        public DateTime? AddedDate { get; set; }
        public string PinCode { get; set; }
        public string OldPinCode { get; set; }
        public DateTime PinCodeExpiryDate { get; set; }
        public bool isValid { get; set; }
        public string Message { get; set; }
        public bool isUpdated { get; set; }


        //public string FirstName { get; set; }
        //public int cpUserType { get; set; }
        //public string Message { get; set; }
        //[Required(ErrorMessage = "* Required")]
        //public bool EmailConfirmed { get; set; }
        ////[Required(ErrorMessage = "* Required")]
        //public string PasswordHash { get; set; }
        //public string SecurityStamp { get; set; }
        //public string PhoneNumber { get; set; }
        //public bool PhoneNumberConfirmed { get; set; }
        //public bool TwoFactorEnabled { get; set; }
        //public DateTime LockoutEndDateUtc { get; set; }
        //public bool LockoutEnabled { get; set; }
        //public int AccessFailedCount { get; set; }
        ////[Required(ErrorMessage = "* Required")]
        //public string UserName { get; set; }
        //public string Discriminator { get; set; }
        //[Required(ErrorMessage = "* Required")]
        //public string FirstName { get; set; }
        //[Required(ErrorMessage = "* Required")]
        //public string LastName { get; set; }
        //[Required(ErrorMessage = "* Required")]
        //public DateTime? DOB { get; set; }
        //public string Gender { get; set; }
        //public string LanguageISOCode { get; set; }
        //public string Address1 { get; set; }
        //public string Address2 { get; set; }
        //[Required(ErrorMessage = "* Required")]
        //public string City { get; set; }
        //public string State { get; set; }
        //public string ZipCode { get; set; }
        ////[Required(ErrorMessage = "* Required")]
        //public string CountryISOCode { get; set; }
        //public string SecondaryEmail { get; set; }
        //public string Fax { get; set; }
        //public string HomePhone { get; set; }
        //public string OfficePhone { get; set; }
        ////[Required(ErrorMessage = "* Required")]
        //public string CellPhone { get; set; }
        //public byte? UserType { get; set; }
        //public bool? Status { get; set; }
        //public string Photo { get; set; }
        //public bool VerifyIPAddress { get; set; }
        //public DateTime DateLastLogin { get; set; }
        //public DateTime DateAdded { get; set; }
        //public DateTime FromDate { get; set; }
        //public DateTime DateTo { get; set; }
        //public string ParentUserID { get; set; }
        //public string PinCode { get; set; }
        //public DateTime PinCodeExpiryDate { get; set; }
        //public string Node { get; set; }
        //public string TimeZoneID { get; set; }
        //public string NationalityCountryISOCode { get; set; }
        //public string BirthCountryISOCode { get; set; }
        //public string Aux1 { get; set; }
        //public string Aux2 { get; set; }
        //public string Aux3 { get; set; }
        //public string Aux4 { get; set; }
        //public string Aux5 { get; set; }
    }
    public class RegisterViewModel
    {
        [Required]
        [EmailAddress]
        [Display(Name = "Email")]
        public string Email { get; set; }

        [Required]
        [StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        [Display(Name = "Password")]
        public string Password { get; set; }

        [DataType(DataType.Password)]
        [Display(Name = "Confirm password")]
        [Compare("Password", ErrorMessage = "The password and confirmation password do not match.")]
        public string ConfirmPassword { get; set; }
        public string UserID { get; set; }

        public string UserName { get; set; }

        public string FullName { get; set; }
        [Required(ErrorMessage = "First name is required")]
        public string FirstName { get; set; }
        [Required(ErrorMessage = "Last name is required")]
        public string LastName { get; set; }
        public DateTime? DateOfBirth { get; set; }
        public string Gender { get; set; }
        //[Required(ErrorMessage = "Address is required")]
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        //[Required(ErrorMessage = "City is required")]
        public string City { get; set; }
        //[Required(ErrorMessage = "State is required")]
        public string State { get; set; }
        //[Required(ErrorMessage = "Post code is required")]
        public string ZipCode { get; set; }
        //[Required(ErrorMessage = "Country is required")]
        public string CountryISOCode { get; set; }
        //[Required(ErrorMessage = "Cell phone is required")]
        public string CellPhone { get; set; }

        public string HomePhone { get; set; }
        public string OfficePhone { get; set; }
        //[Required(ErrorMessage = "Fax is required")]
        public string Fax { get; set; }
        public byte? UserType { get; set; }
        //[Required(ErrorMessage = "Language iso code is required")]
        public string LanguageISOCode { get; set; }
        public bool? Status { get; set; }
        public string Photo { get; set; }
        public bool VerifyIPAddress { get; set; }
        public bool isAdded { get; set; }
        public bool isValid { get; set; }
        [Required(ErrorMessage = "Pin code is required")]
        public string PinCode { get; set; }
        public string Message { get; set; }
        public List<UserMenuViewModel> menus { get; set; }
    }
    public class ResetPasswordViewModel
    {
        [Required]
        public string UserID { get; set; }

        [Required]
        [EmailAddress]
        [Display(Name = "Email")]
        public string Email { get; set; }

        [Required]
        [StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        [Display(Name = "Password")]
        public string Password { get; set; }

        [DataType(DataType.Password)]
        [Display(Name = "Confirm password")]
        [Compare("Password", ErrorMessage = "The password and confirmation password do not match.")]
        public string ConfirmPassword { get; set; }

        public string Code { get; set; }
        public string AKey { get; set; }
        public bool IsSendEmail { get; set; }
        public string FullName { get; set; }
        public bool isValid { get; set; }
        public string Message { get; set; }
    }
    public class Client
    {
        [Key]
        public string Id { get; set; }
        [Required]
        public string Secret { get; set; }
        [Required]
        [MaxLength(100)]
        public string Name { get; set; }
        public ApplicationTypes ApplicationType { get; set; }
        public bool Active { get; set; }
        public int RefreshTokenLifeTime { get; set; }
        [MaxLength(100)]
        public string AllowedOrigin { get; set; }
    }

    public class RefreshToken
    {
        [Key]
        public string Id { get; set; }
        [Required]
        [MaxLength(50)]
        public string Subject { get; set; }
        [Required]
        [MaxLength(50)]
        public string ClientId { get; set; }
        public DateTime IssuedUtc { get; set; }
        public DateTime ExpiresUtc { get; set; }
        [Required]
        public string ProtectedTicket { get; set; }
    }
    public class UsersIPAdressModel
    {
        public string UserID { get; set; }
        [Required(ErrorMessage = "* Required")]
        public string IPAddress { get; set; }
        public bool Status { get; set; }
        public bool VerifyIPAddress { get; set; }
        public string UsersIPAdressID { get; set; }
        public string Message { get; set; }
        public string UserName { get; set; }
    }

    //public class ClientModel
    //{
    //    public static string HostName { get; set; }
    //    public static string DBName { get; set; }
    //    public static string DBPassword { get; set; }
    //    public static string ServerName { get; set; }
    //    public static string DBUser { get; set; }
    //    public static string CompanyLogo { get; set; }
    //    public static string CompanyName { get; set; }
    //    public static string CompanyAddress { get; set; }
    //    public static string CompanyPhone { get; set; }
    //    public static string CompanySubmissions{ get; set; }
    //    public static string IPAddress { get; set; }
    //    public static bool CompanyStatus { get; set; }
    //}
    public class ClientViewModel
    {
        public string HostName { get; set; }
        public string DBName { get; set; }
        public string DBPassword { get; set; }
        public string ServerName { get; set; }
        public string DBUser { get; set; }
        public string CompanyLogo { get; set; }
        public string CompanyName { get; set; }
        public string CompanyAddress { get; set; }
        public string CompanyPhone { get; set; }
        public string CompanySubmissions { get; set; }
        public string IPAddress { get; set; }
        public bool CompanyStatus { get; set; }
    }
}