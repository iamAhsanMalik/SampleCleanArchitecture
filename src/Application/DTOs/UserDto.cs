using Domain.Enums;

namespace Application.DTOs;

public class AgentDto
{
    public string? CustomerID { get; set; }
    public string? Company { get; set; }
    public string? ImageUrl { get; set; }
    public string? UserID { get; set; }
    public string? Email { get; set; }
    public string? Password { get; set; }
    public string? UserName { get; set; }
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    public DateTime? DateOfBirth { get; set; }
    public string? Gender { get; set; }
    public string? FullAddress { get; set; }
    [Required(ErrorMessage = "Address is required")]
    public string? Address1 { get; set; }
    public string? Address2 { get; set; }
    [Required(ErrorMessage = "City is required")]
    public string? City { get; set; }
    [Required(ErrorMessage = "State is required")]
    public string? State { get; set; }
    [Required(ErrorMessage = "Zip code is required")]
    public string? ZipCode { get; set; }
    [Required(ErrorMessage = "Country iso code is required")]
    public string? CountryISOCode { get; set; }
    public string? Phone { get; set; }
    [Required(ErrorMessage = "Cell phone is required")]
    public string? CellPhone { get; set; }
    public string? HomePhone { get; set; }
    public string? OfficePhone { get; set; }
    [Required(ErrorMessage = "Fax is required")]
    public string? Fax { get; set; }
    public byte? UserType { get; set; }
    [Required(ErrorMessage = "Language iso code is required")]
    public string? LanguageISOCode { get; set; }
    public bool? Status { get; set; }
    public string? Photo { get; set; }
    public bool VerifyIPAddress { get; set; }
    public string? AddedBy { get; set; }
    public DateTime? AddedDate { get; set; }
    public string? PinCode { get; set; }
    public string? OldPinCode { get; set; }
    public DateTime PinCodeExpiryDate { get; set; }
    public bool IsValid { get; set; }
    public string? Message { get; set; }
    public bool IsUpdated { get; set; }

}
public class UserDto
{
    public string? ImageUrl { get; set; }
    public string? UserID { get; set; }
    [Required(ErrorMessage = "Email is required")]
    public string? Email { get; set; }
    public string? Password { get; set; }
    public string? UserName { get; set; }
    [Required(ErrorMessage = "First name is required")]
    public string? FirstName { get; set; }
    [Required(ErrorMessage = "Last name is required")]
    public string? LastName { get; set; }
    public DateTime DateOfBirth { get; set; }
    public string? Gender { get; set; }
    public string? FullAddress { get; set; }
    public string? Address1 { get; set; }
    public string? Address2 { get; set; }
    public string? City { get; set; }
    public string? State { get; set; }
    public string? ZipCode { get; set; }
    public string? CountryISOCode { get; set; }
    public string? Phone { get; set; }
    public string? CellPhone { get; set; }
    public string? HomePhone { get; set; }
    public string? OfficePhone { get; set; }
    //[Required(ErrorMessage = "Fax is required")]
    public string? Fax { get; set; }
    public byte? UserType { get; set; }
    //[Required(ErrorMessage = "Language iso code is required")]
    public string? LanguageISOCode { get; set; }
    public bool? Status { get; set; }
    public string? Photo { get; set; }
    public bool VerifyIPAddress { get; set; }
    public string? AddedBy { get; set; }
    public DateTime? AddedDate { get; set; }
    public string? PinCode { get; set; }
    public string? OldPinCode { get; set; }
    public DateTime PinCodeExpiryDate { get; set; }
    public bool IsValid { get; set; }
    public string? Message { get; set; }
    public bool IsUpdated { get; set; }
}
public class RegisterViewModel
{
    [Required, EmailAddress, Display(Name = "Email")]
    public string? Email { get; set; }

    [Required, StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
    [DataType(DataType.Password), Display(Name = "Password")]
    public string? Password { get; set; }

    [DataType(DataType.Password)]
    [Display(Name = "Confirm password")]
    [Compare("Password", ErrorMessage = "The password and confirmation password do not match.")]
    public string? ConfirmPassword { get; set; }
    public string? UserID { get; set; }

    public string? UserName { get; set; }

    public string? FullName { get; set; }
    [Required(ErrorMessage = "First name is required")]
    public string? FirstName { get; set; }
    [Required(ErrorMessage = "Last name is required")]
    public string? LastName { get; set; }
    public DateTime? DateOfBirth { get; set; }
    public string? Gender { get; set; }
    //[Required(ErrorMessage = "Address is required")]
    public string? Address1 { get; set; }
    public string? Address2 { get; set; }
    //[Required(ErrorMessage = "City is required")]
    public string? City { get; set; }
    //[Required(ErrorMessage = "State is required")]
    public string? State { get; set; }
    //[Required(ErrorMessage = "Post code is required")]
    public string? ZipCode { get; set; }
    //[Required(ErrorMessage = "Country is required")]
    public string? CountryISOCode { get; set; }
    //[Required(ErrorMessage = "Cell phone is required")]
    public string? CellPhone { get; set; }

    public string? HomePhone { get; set; }
    public string? OfficePhone { get; set; }
    //[Required(ErrorMessage = "Fax is required")]
    public string? Fax { get; set; }
    public byte? UserType { get; set; }
    //[Required(ErrorMessage = "Language iso code is required")]
    public string? LanguageISOCode { get; set; }
    public bool? Status { get; set; }
    public string? Photo { get; set; }
    public bool VerifyIPAddress { get; set; }
    public bool IsAdded { get; set; }
    public bool IsValid { get; set; }
    [Required(ErrorMessage = "Pin code is required")]
    public string? PinCode { get; set; }
    public string? Message { get; set; }
    public List<UserMenuDto> Menus { get; set; } = new List<UserMenuDto>();
}
public class ResetPasswordDto
{
    [Required]
    public string? UserID { get; set; }

    [Required]
    [EmailAddress]
    [Display(Name = "Email")]
    public string? Email { get; set; }

    [Required]
    [StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
    [DataType(DataType.Password)]
    [Display(Name = "Password")]
    public string? Password { get; set; }

    [DataType(DataType.Password)]
    [Display(Name = "Confirm password")]
    [Compare("Password", ErrorMessage = "The password and confirmation password do not match.")]
    public string? ConfirmPassword { get; set; }
    public string? Code { get; set; }
    public string? AKey { get; set; }
    public bool IsSendEmail { get; set; }
    public string? FullName { get; set; }
    public bool IsValid { get; set; }
    public string? Message { get; set; }
}
public class ClientDto
{
    public string? Id { get; set; }
    [Required]
    public string? Secret { get; set; }
    [Required]
    [MaxLength(100)]
    public string? Name { get; set; }
    public ApplicationTypes ApplicationType { get; set; }
    public bool Active { get; set; }
    public int RefreshTokenLifeTime { get; set; }
    [MaxLength(100)]
    public string? AllowedOrigin { get; set; }
}

public class RefreshToken
{
    public string? Id { get; set; }
    [Required, MaxLength(50)]
    public string? Subject { get; set; }
    [Required, MaxLength(50)]
    public string? ClientId { get; set; }
    public DateTime IssuedUtc { get; set; }
    public DateTime ExpiresUtc { get; set; }
    [Required]
    public string? ProtectedTicket { get; set; }
}
public class UsersIPAdressDto
{
    public string? UserID { get; set; }
    public string? IPAddress { get; set; }
    public bool Status { get; set; }
    public bool VerifyIPAddress { get; set; }
    public string? UsersIPAdressID { get; set; }
    public string? Message { get; set; }
    public string? UserName { get; set; }
}
public class ClientAppConfigDto
{
    public string? HostName { get; set; }
    public string? DBName { get; set; }
    public string? DBPassword { get; set; }
    public string? ServerName { get; set; }
    public string? DBUser { get; set; }
    public string? CompanyLogo { get; set; }
    public string? CompanyName { get; set; }
    public string? CompanyAddress { get; set; }
    public string? CompanyPhone { get; set; }
    public string? CompanySubmissions { get; set; }
    public string? IPAddress { get; set; }
    public bool CompanyStatus { get; set; }
}