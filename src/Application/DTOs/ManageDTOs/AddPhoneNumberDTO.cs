namespace Application.DTOs.ManageDTOs;

public class AddPhoneNumberDTO
{
    [Required]
    [Phone]
    [Display(Name = "Phone number")]
    public string? PhoneNumber { get; set; }
}
