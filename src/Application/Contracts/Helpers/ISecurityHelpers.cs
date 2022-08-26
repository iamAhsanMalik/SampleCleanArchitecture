namespace Application.Contracts.Helpers;

public interface ISecurityHelpers
{
    string Decrypt(string inputText);
    string DecryptQueryString(string inputText);
    string Encrypt(string inputText);
    string EncryptQueryString(string inputText);
}