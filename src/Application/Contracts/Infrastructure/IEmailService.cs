using Application.Models.EmailModels;
using Mailjet.Client.TransactionalEmails.Response;

namespace Application.Contracts.Infrastructure;
public interface IEmailService
{
    Task<TransactionalEmailResponse> SendEmailByMailJetAsync(MailJetEmailRequest mailJetEmailRequest);
    Task<string> SendEmailbySMTPAsync(SMTPEmailRequest smtpEmailRequest);
}