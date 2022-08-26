using System.ComponentModel.DataAnnotations;

namespace Application.DTOs.AccountDTOs;

public class ExternalLoginConfirmationDTO
{
    [Required]
    [EmailAddress]
    public string? Email { get; set; }
}
