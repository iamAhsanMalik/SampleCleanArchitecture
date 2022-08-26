namespace Application.Models.EmailModels;
public class MailJetEmailRequest
{
    public string? EmailSubject { get; set; }
    public string? EmailBody { get; set; }
    public string? ToEmail { get; set; }
    public string? FromEmail { get; set; }
}
