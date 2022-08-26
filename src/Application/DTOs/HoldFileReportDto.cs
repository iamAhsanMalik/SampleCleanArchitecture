namespace Application.DTOs;
public class HoldFileReportDto
{
    public string? Date { get; set; }
    public string? InvoiceNo { get; set; }
    public string? Agent { get; set; }
    public decimal AmountInFC { get; set; }
    public decimal AmountInLC { get; set; }
    public decimal TotalFCBalance { get; set; }
    public decimal TotalLCBalance { get; set; }
    public string? TempTable { get; set; }
    public bool Isvalid { get; set; }
    public DateTime FromDate { get; set; }
    public DateTime ToDate { get; set; }
    public string? VoucherNo { get; set; }
    public string? Memo { get; set; }
    public string? TransactionID { get; set; }
    public string? Recipient { get; set; }
    public string? AgentPrefix { get; set; }
    public string? SenderName { get; set; }
    public string? CustomerId { get; set; }
    public string? Description { get; set; }
    public string? PaymentNo { get; set; }
    public string? ReceiverName { get; set; }
    public string? AccountCode { get; set; }
}

