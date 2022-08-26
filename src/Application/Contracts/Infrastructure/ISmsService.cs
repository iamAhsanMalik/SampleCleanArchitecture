namespace Application.Contracts.Infrastructure;

public interface ISmsService
{
    Task SMSSenderAsync(string number, string message);
}
