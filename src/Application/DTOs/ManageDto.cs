namespace Application.DTOs;

public class IndexDto
{
    public bool HasPassword { get; set; }
    public IList<UserLoginInfo> Logins { get; set; } = new List<UserLoginInfo>();
    public string? PhoneNumber { get; set; }
    public bool TwoFactor { get; set; }
    public bool BrowserRemembered { get; set; }
}

public class ManageLoginsDto<T>
{
    public IList<UserLoginInfo> CurrentLogins { get; set; } = new List<UserLoginInfo>();
    public IList<T> OtherLogins { get; set; } = new List<T>();
}

public class FactorViewModel
{
    public string? Purpose { get; set; }
}

public class SetPasswordDto
{
    [Required]
    [StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
    [DataType(DataType.Password)]
    [Display(Name = "New password")]
    public string? NewPassword { get; set; }

    [DataType(DataType.Password)]
    [Display(Name = "Confirm new password")]
    [Compare("NewPassword", ErrorMessage = "The new password and confirmation password do not match.")]
    public string? ConfirmPassword { get; set; }
}

public class ChangePasswordDto
{
    [Required]
    [DataType(DataType.Password)]
    [Display(Name = "Current password")]
    public string? OldPassword { get; set; }

    [Required]
    [StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
    [DataType(DataType.Password)]
    [Display(Name = "New password")]
    public string? NewPassword { get; set; }

    [DataType(DataType.Password)]
    [Display(Name = "Confirm new password")]
    [Compare("NewPassword", ErrorMessage = "The new password and confirmation password do not match.")]
    public string? ConfirmPassword { get; set; }
}

public class AddPhoneNumberDto
{
    [Required,Phone,Display(Name = "Phone Number")]
    public string? Number { get; set; }
}

public class VerifyPhoneNumberDto
{
    [Required, Display(Name = "Code")]
    public string? Code { get; set; }

    [Required, Phone, Display(Name = "Phone Number")]
    public string? PhoneNumber { get; set; }
}

public class ConfigureTwoFactorDto<T>
{
    public string? SelectedProvider { get; set; }
    public ICollection<T>? Providers { get; set; }
}