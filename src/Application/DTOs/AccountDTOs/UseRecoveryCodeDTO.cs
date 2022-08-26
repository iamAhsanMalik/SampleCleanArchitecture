namespace Application.DTOs.AccountDTOs;

public class UseRecoveryCodeDTO
{
    [Required]
    public string? Code { get; set; }

    public string? ReturnUrl { get; set; }
}
