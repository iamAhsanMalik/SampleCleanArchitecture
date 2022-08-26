namespace Application.DTOs.ManageDTOs;

public class IndexDTO<IdentityUserLoginInfo>
{
    public bool HasPassword { get; set; }

    public IList<IdentityUserLoginInfo>? Logins { get; set; }

    public string? PhoneNumber { get; set; }

    public bool TwoFactor { get; set; }

    public bool BrowserRemembered { get; set; }

    public string? AuthenticatorKey { get; set; }
}
