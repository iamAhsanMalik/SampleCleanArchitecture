using System.ComponentModel.DataAnnotations;

namespace Application.DTOs.AccountDTOs;

public class ForgotPasswordDTO
{
    [Required]
    [EmailAddress]
    public string? Email { get; set; }
}
