
namespace Infrastructure.Services.Communication.Email;

internal class EmailServices : IEmailService
{
    private readonly MailJetConfig _mailJetConfig;
    public EmailServices(IOptions<MailJetConfig> mailJetConfig)
    {
        _mailJetConfig = mailJetConfig.Value;
    }
    public async Task SendEmailAsync(string toEmail, string subject, string message)
    {
        // Plug in your email service here to send an email.

        if (string.IsNullOrEmpty(_mailJetConfig.ApiKey) || string.IsNullOrEmpty(_mailJetConfig.SecretKey))
        {
            throw new Exception("Missing MailJet configurations ");
        }

        await Execute(_mailJetConfig, subject, message, toEmail);
    }

    public async Task<Mailjet.Client.TransactionalEmails.Response.TransactionalEmailResponse> Execute(MailJetConfig config, string subject, string message, string toEmail)
    {

        MailjetClient _client = new MailjetClient(config.ApiKey, config.SecretKey);

        // construct your email with builder
        var email = new TransactionalEmailBuilder()
               .WithFrom(new SendContact("undeclaredvariable@protonmail.com"))
               .WithSubject(subject)
               .WithHtmlPart(message)
               .WithTo(new SendContact(toEmail))
               .Build();
        // invoke API to send email
        return await _client.SendTransactionalEmailAsync(email);
    }
}
