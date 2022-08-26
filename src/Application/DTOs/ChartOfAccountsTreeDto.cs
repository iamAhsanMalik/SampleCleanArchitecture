namespace Application.DTOs;

public class ChartOfAccountsTreeDto
{
    public int AccountID { get; set; }
    public int AccountCode { get; set; }
    public int ParentAccountCode { get; set; }
    public string? AccountName { get; set; }
    public bool AccountStatus { get; set; }
    public string? CurrencyISOCode { get; set; }
    public string? Prefix { get; set; }
    public DateTime AddedDate { get; set; }
    public decimal ForeignCurrencyBalance { get; set; }
    public decimal LocalCurrencyBalance { get; set; }
    public bool IsHead { get; set; }

}