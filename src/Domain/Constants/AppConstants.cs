namespace Domain.Constants;

public static class AppConstants
{
    public static string AllowedUserNameCharacters
    {
        get => "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.@+";
    }

}
#region Image Type
public static class ImageType
{
    public static string PNG { get => "*.png"; }
    public static string JPG { get => "*.jpg"; }
    public static string JPEG { get => "*.jpeg"; }
    public static string Gif { get => "*.gif"; }
}
#endregion

#region Database Connection String Properties
public static class ConnectionString
{
    public static string ConStr { get => "ConStr"; }
    public static string ConStrOld { get => "ConStrOld"; }
    public static string EmailConStr { get => "EmailConStr"; }
    public static string ConStrAA { get => "ConStrAA"; }
    public static string ConStrEmailer { get => "ConStrEmiler"; }
    public static string BlogConstr { get => "mySqlBlogConstr"; }
    public static string RadioTVConstr { get => "mySqlRadioTVConstr"; }
    public static string TVConstr { get => "mySqlTVConstr"; }
    public static string WebPagesConstr { get => "myWebPagesConstr"; }
}
#endregion

#region Get Month by Id
public class Month
{
    public string? January { get => "January"; }
    public string? February { get => "February"; }
    public string? March { get => "March"; }
    public string? April { get => "April"; }
    public string? May { get => "May"; }
    public string? June { get => "June"; }
    public string? July { get => "July"; }
    public string? August { get => "August"; }
    public string? September { get => "September"; }
    public string? October { get => "October"; }
    public string? November { get => "November"; }
    public string? December { get => "December"; }
}
#endregion