namespace Application.DTOs;

public class TrialbalanceDto
{
    public string? Table { get; set; }
    public bool ShowAll { get; set; }
    public bool HaveData { get; set; }
    public DateTime DateFrom { get; set; }
    public DateTime DateTo { get; set; }
    public DateTime FromDate { get; set; }
    public DateTime ToDate { get; set; }
    public string? AccountName { get; set; }
    public string? SubName3 { get; set; }
    public string? SubName2 { get; set; }
    public string? SubName1 { get; set; }
    public string? AccountHead { get; set; }
    public string? Debit { get; set; }
    public string? Credit { get; set; }
    public string? AccountCode { get; set; }
    public string? LocalCurrencyBalance { get; set; }
}