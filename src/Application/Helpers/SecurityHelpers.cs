using Application.Contracts.Helpers;
using System.Security.Cryptography;

namespace Application.Helpers;

internal class SecurityHelpers : ISecurityHelpers
{
    private const string _eNCRYPTION_KEY = "#i%R&e^M~*";

    /// <summary>
    /// The salt value used to strengthen the encryption.
    /// </summary>
    private readonly byte[] _sALT = Encoding.ASCII.GetBytes(_eNCRYPTION_KEY.Length.ToString());

    public string Encrypt(string inputText)
    {
        try
        {
            Aes aesCipher = Aes.Create();
            byte[] plainText = Encoding.Unicode.GetBytes(inputText);
            PasswordDeriveBytes SecretKey = new PasswordDeriveBytes(_eNCRYPTION_KEY, _sALT);

            using (ICryptoTransform encryptor = aesCipher.CreateEncryptor(SecretKey.GetBytes(32), SecretKey.GetBytes(16)))
            using (MemoryStream memoryStream = new MemoryStream())
            using (CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
            {
                cryptoStream.Write(plainText, 0, plainText.Length);
                cryptoStream.FlushFinalBlock();
                return Convert.ToBase64String(memoryStream.ToArray());
            }
        }
        catch (Exception)
        {
            return "";
        }
    }

    public string Decrypt(string inputText)
    {
        try
        {
            Aes aesCipher = Aes.Create();

            byte[] encryptedData = Convert.FromBase64String(inputText);
            PasswordDeriveBytes secretKey = new PasswordDeriveBytes(_eNCRYPTION_KEY, _sALT);

            using (ICryptoTransform decryptor = aesCipher.CreateDecryptor(secretKey.GetBytes(32), secretKey.GetBytes(16)))
            using (MemoryStream memoryStream = new MemoryStream(encryptedData))
            using (CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read))
            {
                byte[] plainText = new byte[encryptedData.Length];
                int decryptedCount = cryptoStream.Read(plainText, 0, plainText.Length);
                return Encoding.Unicode.GetString(plainText, 0, decryptedCount);
            }
        }
        catch (Exception)
        {
            return "";
        }
    }

    public string EncryptQueryString(string inputText)
    {
        byte[] plainText = Encoding.UTF8.GetBytes(inputText);

        using (Aes aesCipher = Aes.Create())
        {
            PasswordDeriveBytes secretKey = new PasswordDeriveBytes(_eNCRYPTION_KEY, _sALT);
            using (ICryptoTransform encryptor = aesCipher.CreateEncryptor(secretKey.GetBytes(32), secretKey.GetBytes(16)))
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                    {
                        cryptoStream.Write(plainText, 0, plainText.Length);
                        cryptoStream.FlushFinalBlock();
                        string base64 = Convert.ToBase64String(memoryStream.ToArray());

                        // Generate a string that won't get screwed up when passed as a query string.
                        return System.Web.HttpUtility.UrlEncode(base64);
                    }
                }
            }
        }
    }

    public string DecryptQueryString(string inputText)
    {
        byte[] encryptedData = Convert.FromBase64String(inputText);
        PasswordDeriveBytes secretKey = new PasswordDeriveBytes(_eNCRYPTION_KEY, _sALT);
        using (Aes aesCipher = Aes.Create())
        {
            using (ICryptoTransform decryptor = aesCipher.CreateDecryptor(secretKey.GetBytes(32), secretKey.GetBytes(16)))
            {
                using (MemoryStream memoryStream = new MemoryStream(encryptedData))
                {
                    using (CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read))
                    {
                        byte[] plainText = new byte[encryptedData.Length];
                        cryptoStream.Read(plainText, 0, plainText.Length);
                        return Encoding.UTF8.GetString(plainText);
                    }
                }
            }
        }
    }
}
