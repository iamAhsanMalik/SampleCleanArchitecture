namespace Application.DTOs;
public class AgentReportDto
{
    public DateTime FromDate { get; set; }
    public DateTime ToDate { get; set; }
    public string? AccountCode { get; set; }
    public string? Country { get; set; }

    public int TotalTransactions { get; set; }
    public bool Isvalid { get; set; }
    public decimal AdminFCAmount { get; set; }
    public decimal AdminLCAmount { get; set; }
    public decimal AgentFCAmount { get; set; }
    public decimal AgentLCAmount { get; set; }
    public decimal FXMarginFCAmount { get; set; }
    public decimal FXMarginLCAmount { get; set; }
    public string? HtmlTable { get; set; }
    public bool ShowAll { get; set; }
    public decimal TotalFCBalance { get; set; }
    public decimal TotalLCBalance { get; set; }
    public string? TempTable { get; set; }
    public int AccountID { get; set; }
    public string? AccountName { get; set; }
    public string? ErrMessage { get; set; }
    public bool IsPDF { get; set; }
}
