using Application.Models.EmailModels;
using Mailjet.Client.TransactionalEmails.Response;

namespace Application.Contracts.Infrastructure;
public interface IEmailService
{
    Task<TransactionalEmailResponse> MailJetEmailSenderAsync(MailJetEmailRequest mailJetEmailRequest);
    Task<string> SMTPEmailSenderAsync(SMTPEmailRequest smtpEmailRequest);
}