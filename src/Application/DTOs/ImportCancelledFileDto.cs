namespace Application.DTOs;
public class ImportCancelledFileDto
{
    public int Id { get; set; }
    public bool IsError { get; set; }
    public bool IsFileImported { get; set; }
    public bool IsRowValid { get; set; }
    public string? RowErrMessage { get; set; }
    public bool IsPosted { get; set; }
    public string? CpUserRefNo { get; set; }
    public string? Sheet { get; set; }
    public string? ErrMessage { get; set; }
    public bool HasHeader { get; set; }
    public string? ExcelFilePath { get; set; }
    public int RowNumber { get; set; }
    public DateTime TransactionDate { get; set; }
    public string? AgentName { get; set; }
    public decimal Rate { get; set; }
    public string? PayoutCurrency { get; set; }
    public decimal FCAmount { get; set; }
    public decimal Payinamount { get; set; }
    public decimal AdminCharges { get; set; }
    public decimal AgentCharges { get; set; }
    public decimal Charges { get; set; }
    public int TRANID { get; set; }
    public int CustomerId { get; set; }
    public string? CustomerName { get; set; }
    public string? FatherName { get; set; }
    public string? Recipient { get; set; }
    public string? Phone { get; set; }
    public string? Address { get; set; }
    public string? City { get; set; }
    public string? Code { get; set; }
    public string? PaymentNo { get; set; }
    public string? BenePaymentMethod { get; set; }
    public string? SendingCountry { get; set; }
    public string? ReceivingCountry { get; set; }
    public string? Status { get; set; }
    public DateTime PostingDate { get; set; }

    //public List<Grid> grid = new List<Grid>();

    public List<SheetsDto> sheets = new List<SheetsDto>();
}
