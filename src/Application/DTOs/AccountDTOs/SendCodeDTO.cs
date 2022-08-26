namespace Application.DTOs.AccountDTOs;

public class SendCodeDTO<Authenticators>
{
    public string? SelectedProvider { get; set; }

    public ICollection<Authenticators>? Providers { get; set; }

    public string? ReturnUrl { get; set; }

    public bool RememberMe { get; set; }
}
