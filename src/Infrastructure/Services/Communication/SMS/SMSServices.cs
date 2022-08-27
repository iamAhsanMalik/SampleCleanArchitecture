namespace Infrastructure.Services.Communication.SMS;

internal class SMSServices : ISmsService
{
    public Task SendSmsByTwillioAsync(string number, string message)
    {
        // Plug in your SMS service here to send a text message.
        return Task.FromResult(0);
    }
}
