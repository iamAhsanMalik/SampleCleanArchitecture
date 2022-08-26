namespace Application.Configurations.SMTP;

public class SMTPConfig
{
    public string? SmtpHost { get; set; }
    public int SmtpPort { get; set; }
    public string? SenderEmail { get; set; }
    public string? Account { get; set; }
    public string? Password { get; set; }
    public bool SecureSocketOptions { get; set; }
}

