namespace Application.DTOs.ManageDTOs;

public class ConfigureTwoFactorDTO<Authenticators>
{
    public string? SelectedProvider { get; set; }

    public ICollection<Authenticators>? Providers { get; set; }
}
