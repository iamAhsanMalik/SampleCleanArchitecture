using Application.Contracts.Helpers;

namespace Application.Helpers;
internal class ValidatorHelpers : IValidatorHelpers
{
    public bool EmailValidator(string email)
    {
        //regular expression pattern for valid email
        //addresses, allows for the following domains:
        //com,edu,info,gov,int,mil,net,org,biz,name,museum,coop,aero,pro,tv
        string pattern = @"^[-a-zA-Z0-9][-.a-zA-Z0-9]*@[-.a-zA-Z0-9]+(\.[-.a-zA-Z0-9]+)*\.(com|edu|info|gov|int|mil|net|org|biz|name|museum|coop|aero|pro|tv|[a-zA-Z]{2})$";
        //Regular expression object
        Regex check = new Regex(pattern, RegexOptions.IgnorePatternWhitespace);
        //boolean variable to return to calling method
        bool valid = false;

        //make sure an email address was provided
        if (string.IsNullOrEmpty(email))
        {
            valid = false;
        }
        else
        {
            //use IsMatch to validate the address
            valid = check.IsMatch(email);
        }
        //return the value to the calling method
        return valid;
    }

}
