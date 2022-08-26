namespace Application.DTOs;

public class ChartOfAccountsDto
{
    public int AccountID { get; set; }
    public string? AccountCode { get; set; }
    public string? AccountName { get; set; }
    public string? ParentAccountCode { get; set; }
    public string? ParentAccountName { get; set; }
    public string? AccountCodeMask { get; set; }
    public bool AccountStatus { get; set; }
    public string? AccountDescription { get; set; }
    public string? AddedBy { get; set; }
    public DateTime AddedDate { get; set; }
    public int ClassID { get; set; }
    public int TypeID { get; set; }
    public string? Prefix { get; set; }
    public string? CurrencyISOCode { get; set; }
    //[Required]
    public string? CountryISONumericCode { get; set; }
    public bool RevalueYN { get; set; }
    public decimal Revalue { get; set; }
    public decimal DebiteLimit { get; set; }
    public decimal CreditLimit { get; set; }
    public decimal TaxPercentage { get; set; }
    public decimal VATWHT { get; set; }
    public bool Charges { get; set; }
    public bool IsHead { get; set; }
    public DateTime FromDate { get; set; }
    public DateTime ToDate { get; set; }
    public bool Isvalid { get; set; }
    public decimal ForeignCurrencyBalance { get; set; }
    public decimal LocalCurrencyBalance { get; set; }
    public string? HtmlTable { get; set; }
    public bool ShowAll { get; set; }
    public decimal TotalBalance { get; set; }
    public string? TempTable { get; set; }
    public bool IsPrint { get; set; }
    public bool IsVisible { get; set; }

    public int PreAccountID { get; set; }
    public string? PreAccountCode { get; set; }
    public string? PreParentAccountCode { get; set; }
    public string? PreParentAccountName { get; set; }
    public string? PinCode { get; set; }
    public string? AccountTitle { get; set; }
}