namespace Application.DTOs;


public class UpdateBuyerDto
{
    public bool IsHide { get; set; }
    public string? CancellationMessage { get; set; }
    public bool HideContent { get; set; }
    public string? CpGLReferenceID { get; set; }
    public string? TransactionID { get; set; }
    public string? CustomerTransactionID { get; set; }
    public bool IsPosted { get; set; }
    public string? ListTable { get; set; }
    public string? VendorId { get; set; }
    public string? ErrMessage { get; set; }
    public string? InvoiceNo { get; set; }
    public DateTime Date { get; set; }
    public string? HtmlTable1 { get; set; }
    public string? Status { get; set; }
    public decimal FC_Amount { get; set; }
    public decimal Rate { get; set; }
    public decimal Pounds { get; set; }
    public DateTime PaidDate { get; set; }
    public string? CustomerPrefix { get; set; }
    public string? SenderName { get; set; }
    public string? Phone { get; set; }
    public string? AccountName { get; set; }
    public string? PaymentNo { get; set; }
    public string? VoucherNo { get; set; }
    public string? HtmlTable { get; set; }
    public string? PaidGLReferenceID { get; set; }
    public string? CurrencyISOCode { get; set; }
    public string? AgentCurrRate { get; set; }
    public string? CurrencyType { get; set; }

    public string? HoldAccount { get; set; }
    public string? PinCode { get; set; }

    public string? BuyerPrefix { get; set; }
}