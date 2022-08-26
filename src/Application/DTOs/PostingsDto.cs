namespace Application.DTOs;

public class PostingsDto
{
    public bool IsPosted { get; set; }
    public bool IsListPosted { get; set; }
    public bool IsBtnShow { get; set; }
    public string? AccountCode { get; set; }
    public string? ErrMessage { get; set; }
    public string? HtmlTable { get; set; }
    public string? ListTable { get; set; }
    public string? CpRefID { get; set; }
    public string? CpEntryID { get; set; }
    public string? CpUserRefNo { get; set; }
    public string? DrCr { get; set; }
    public string? Table { get; set; }
    public string? CurrencyISOCode { get; set; }
    public int GLReferenceID { get; set; }
    public string? TypeiCode { get; set; }
    public string? ReferenceNo { get; set; }
    public string? Memo { get; set; }
    public string? PostedBy { get; set; }
    public DateTime DatePosted { get; set; }
    public string? AuthorizedBy { get; set; }
    public DateTime AuthorizedDate { get; set; }
    public string? CpActionType { get; set; }
    public string? VoucherNumber { get; set; }
    public string? PinCode { get; set; }
    public string? HtmlTable1 { get; set; }
    public string? StartingRange { get; set; }
    public string? EndingRange { get; set; }
    public DateTime FromDate { get; set; }
    public DateTime ToDate { get; set; }
    public string? Attachments { get; set; }

    public List<GLTransactionDto> gLTransactionList = new List<GLTransactionDto>();
    public List<AddUpdateDto> pAddUpdateViewModel = new List<AddUpdateDto>();
    public string? Payee { get; set; }
    public string? Customer { get; set; }
    public string? Type { get; set; }
}
public class GLTransactionDto
{
    #region GlTransactions
    public int TransactionID { get; set; }
    public int GLReferenceID { get; set; }
    public string? Type { get; set; }
    public string? AccountTypeiCode { get; set; }
    public string? Memo { get; set; }
    public decimal BaseAmount { get; set; }
    public decimal LocalAmount { get; set; }
    public string? ForeignCurrencyISOCode { get; set; }
    public decimal ForeignCurrencyAmount { get; set; }
    public decimal LocalCurrencyAmount { get; set; }
    public decimal ExchangeRate { get; set; }
    public string? GLPersonTypeiCode { get; set; }
    public string? PersonID { get; set; }
    public short DimensionId { get; set; }
    public short Dimension2Id { get; set; }
    public string? AddedBy { get; set; }
    public DateTime AddedDate { get; set; }
    public string? AccountCode { get; set; }
    public string? LocalCurrencyBalance { get; set; }
    #endregion
}

public class AccountBalanceDto
{
    public string? AccountCode { get; set; }
    public decimal ForeignCurrencyBalance { get; set; }
    public decimal LocalCurrencyBalance { get; set; }
}

public class AccountDetailsDto
{
    public string? AccountCode { get; set; }
    public string? LocalCurrencyCode { get; set; }
    public string? ForeignCurrencyCode { get; set; }
    public decimal CreditMaxRate { get; set; }
    public decimal CreditMinRate { get; set; }
    public decimal ForeignCurrencyBalance { get; set; }
    public decimal LocalCurrencyBalance { get; set; }
    public decimal CurrencyRate { get; set; }
    public string? CurrencyType { get; set; }
}

public class AddUpdateDto
{
    public string? BankCashAccountCode { get; set; }
    public string? GeneralAccountCode { get; set; }
    public decimal LocalCurrencyAmount { get; set; }
    public decimal ForeignCurrencyAmount { get; set; }
    public decimal ExchangeRate { get; set; }
    public string? DrCr { get; set; }
    public string? CpUserRefNo { get; set; }
    public int CpEntryID { get; set; }
    public decimal CreditMaxRate { get; set; }
    public decimal CreditMinRate { get; set; }
    public string? Memo { get; set; }
    public string? CurrencyISoCode { get; set; }
}
public class ReturnDto
{
    public bool IsPosted { get; set; }
    public string? ErrMessage { get; set; }
    public string? HtmlTable { get; set; }
}

public class EditDto
{
    public string? CpEntryID { get; set; }
    public string? AccountCode { get; set; }
    public decimal ForeignCurrencyAmount { get; set; }
    public decimal LocalCurrencyAmount { get; set; }
    public string? Memo { get; set; }
    public decimal CreditMaxRate { get; set; }
    public decimal CreditMinRate { get; set; }
    public string? DrCr { get; set; }
    public string? ForeignCurrencyISOCode { get; set; }
    public string? LocalCurrencyISOCode { get; set; }
    public decimal CurrencyRate { get; set; }
    public string? CurrencyType { get; set; }
    public decimal ForeignCurrencyBalance { get; set; }
    public decimal LocalCurrencyBalance { get; set; }


}
public class DeleteDto
{
    public bool IsPosted { get; set; }
    public string? HtmlTable { get; set; }
    public string? CpUserRefNo { get; set; }
    public string? Attachments { get; set; }
}