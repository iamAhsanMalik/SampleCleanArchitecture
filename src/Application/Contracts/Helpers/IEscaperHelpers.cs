namespace Application.Contracts.Helpers;
public interface IEscaperHelpers
{
    string DecodePath(string token);
    string DisplaySecureImage(string fullPath);
    string EncodePath(string fullPath);
    string GetContentType(string extension);
    Regex GetEscaperChar();
    string GetPlainUrl(string fullPath);
    string GetPublicUrl(string fullPath, string fileName);
    string GetSecureUrl(string fullPath, string fileName);
    Regex GetUnescaperChar();
    bool ValidateUrlExpiryDate(DateTime sentDate);
}