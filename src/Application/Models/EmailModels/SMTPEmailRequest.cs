namespace Application.Models.EmailModels;
public class SMTPEmailRequest
{
    public string EmailSubject { get; set; } = null!;
    public string EmailBody { get; set; } = null!;
    public string ToEmail { get; set; } = null!;
    public string FromEmail { get; set; } = string.Empty;
    public ICollection<IFormFile>? Attachments { get; set; }
}
