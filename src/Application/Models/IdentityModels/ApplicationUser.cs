namespace Application.Models.IdentityModels;

public class ApplicationUser : IdentityUser
{
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    //public string? AiuUserId { get; set; }
}
