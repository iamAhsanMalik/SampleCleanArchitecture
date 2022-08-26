using STM.AIU.Application.Contracts.Infrastructure;

namespace Infrastructure.Services.Communication.SMS;

internal class SMSService : ISmsService
{
    public Task SMSSenderAsync(string number, string message)
    {
        // Plug in your SMS service here to send a text message.
        return Task.FromResult(0);
    }
}
