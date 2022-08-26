namespace Application.DTOs;
public class JournalInquiryDto
{
    public DateTime FromDate { get; set; }
    public DateTime ToDate { get; set; }
    public DateTime ExcelFromDate { get; set; }
    public DateTime ExcelToDate { get; set; }
    public DateTime PdfFromDate { get; set; }
    public DateTime PdfToDate { get; set; }
    public bool ShowAll { get; set; }
    public bool ExcelShowAll { get; set; }
    public bool PdfShowAll { get; set; }
    public string? Memo { get; set; }
    public string? VoucherNumber { get; set; }
    public string? AccountTypeICode { get; set; }
    public string? Table { get; set; }
    public string? ErrMessage { get; set; }
    public bool HaveData { get; set; }
    public string? AddedDate { get; set; }
    public string? TypeiCode { get; set; }
    public string? AuthBy { get; set; }
    public string? PostedBy { get; set; }
    public decimal LocalCurrencyAmount { get; set; }
    public decimal ForeignCurrencyAmount { get; set; }
    public string? ForeignCurrencyISOCode { get; set; }
    public string? AccountTitle { get; set; }
    public string? GLReferenceID { get; set; }
    public string? AccountCode { get; set; }
    public string? StartDate { get; set; }
    public string? EndDate { get; set; }
}
