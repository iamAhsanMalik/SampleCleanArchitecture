namespace Application.Helpers;
internal class GeneratorHelpers : IGeneratorHelpers
{
    public string GenerateRandomCode(int codeCount = 5)
    {
        string allChar = "1,2,3,4,5,6,7,8,9,0";
        string[] allCharArray = allChar.Split(',');
        string randomCode = "";
        int temp = -1;

        Random rand = new Random();
        for (int i = 0; i < codeCount; i++)
        {
            if (temp != -1)
            {
                rand = new Random(i * temp * (int)DateTime.Now.Ticks);
            }
            int t = rand.Next(9);
            if (temp != -1 && temp == t)
            {
                return GenerateRandomCode(codeCount);
            }
            temp = t;
            randomCode += allCharArray[t];
        }
        return randomCode;
    }

    public string GenerateRandomDigit(int codeCount = 5)
    {
        string allChar = "1,2,3,4,5,6,7,8,9,0";
        string[] allCharArray = allChar.Split(',');
        string randomCode = "";
        int temp = -1;

        Random rand = new Random();
        for (int i = 0; i < codeCount; i++)
        {
            if (temp != -1)
            {
                rand = new Random(i * temp * (int)DateTime.Now.Ticks);
            }
            int t = rand.Next(9);
            if (temp != -1 && temp == t)
            {
                return GenerateRandomDigit(codeCount);
            }
            temp = t;
            randomCode += allCharArray[t];
        }
        return randomCode;
    }

    public string GenerateRandomString(int length = 5)
    {
        string allowedLetterChars = "abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ";
        string allowedNumberChars = "23456789";
        char[] chars = new char[length];
        Random rd = new Random();

        bool useLetter = true;
        for (int i = 0; i < length; i++)
        {
            if (useLetter)
            {
                chars[i] = allowedLetterChars[rd.Next(0, allowedLetterChars.Length)];
                useLetter = false;
            }
            else
            {
                chars[i] = allowedNumberChars[rd.Next(0, allowedNumberChars.Length)];
                useLetter = true;
            }
        }

        return new string(chars);
    }
}
