namespace Application.DTOs.AccountDTOs;

public class VerifyAuthenticatorCodeDTO
{
    [Required]
    public string? Code { get; set; }

    public string? ReturnUrl { get; set; }

    [Display(Name = "Remember this browser?")]
    public bool RememberBrowser { get; set; }

    [Display(Name = "Remember me?")]
    public bool RememberMe { get; set; }
}
