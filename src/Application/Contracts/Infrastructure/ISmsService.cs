namespace Application.Contracts.Infrastructure;

public interface ISmsService
{
    Task SendSmsByTwillioAsync(string number, string message);
}
