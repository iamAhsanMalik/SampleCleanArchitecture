namespace Application.DTOs;

public class BuyerReportDto
{
    public DateTime FromDate { get; set; }
    public DateTime ToDate { get; set; }
    public string? AccountCode { get; set; }
    public string? OpeningBalanceDate { get; set; }
    public DateTime TransactionDate { get; set; }
    public int TotalTransactions { get; set; }
    public bool Isvalid { get; set; }
    public decimal DebitFCTotal { get; set; }
    public decimal DebitLCTotal { get; set; }
    public decimal CreditFCTotal { get; set; }
    public decimal CreditLCTotal { get; set; }
    public string? HtmlTable { get; set; }
    public bool ShowAll { get; set; }
    public decimal TotalFCBalance { get; set; }
    public decimal TotalLCBalance { get; set; }
    public string? TempTable { get; set; }
    public int AccountID { get; set; }
    public string? AccountName { get; set; }
    public string? ErrMessage { get; set; }
    public bool IsPDF { get; set; }
    public DateTime AddedDate { get; set; }
    public string? PDate { get; set; }
    public int Total { get; set; }
    public string? PTotal { get; set; }
    public decimal LocalCurrencyBalance { get; set; }
    public decimal ForeignCurrencyBalance { get; set; }
    public string? VoucherNumber { get; set; }
    public int SortOrder { get; set; }
}
