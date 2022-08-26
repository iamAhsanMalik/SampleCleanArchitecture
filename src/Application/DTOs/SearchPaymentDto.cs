namespace Application.DTOs;

public class SearchPaymentDto
{
    public DateTime PostingDate { get; set; }
    public string? AgentPrefix { get; set; }
    public string? HoldAccount { get; set; }
    public DateTime CancelledDate { get; set; }
    public DateTime PaidDate { get; set; }
    public DateTime HoldDate { get; set; }
    public string? CustomerPrefix { get; set; }
    public string? CancellationReason { get; set; }
    public string? AccountName { get; set; }
    public string? VoucherNo { get; set; }
    public string? InvoiceNo { get; set; }
    public string? PaymentNo { get; set; }
    public string? CustomerTransactionID { get; set; }
    public string? RecevingCountry { get; set; }
    public string? SenderName { get; set; }
    public string? Phone { get; set; }
    public string? Beneficiary { get; set; }
    public string? BenePhone { get; set; }
    public string? PayInAmount { get; set; }
    public string? PayOutAmount { get; set; }
    public DateTime FromDate { get; set; }
    public string? CustomerID { get; set; }
    public string? TransactionID { get; set; }
    public string? Currency { get; set; }
    public DateTime Date { get; set; }
    public string? Recipient { get; set; }
    public string? Status { get; set; }
    public decimal FC_Amount { get; set; }
    public decimal Rate { get; set; }
    public decimal Pounds { get; set; }
}