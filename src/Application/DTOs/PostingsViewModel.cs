using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class PostingViewModel
    {
        public bool isPosted { get; set; }
        public bool isListPosted { get; set; }
        public bool isBtnShow { get; set; }
        public string AccountCode { get; set; }
        public string ErrMessage { get; set; }
        public string HtmlTable { get; set; }
        public string ListTable { get; set; }
        public string cpRefID { get; set; }
        public string cpEntryID { get; set; }
        public string cpUserRefNo { get; set; }
        public string DrCr { get; set; }
        public string table { get; set; }
        public string CurrencyISOCode { get; set; }
        public int GLReferenceID { get; set; }
        public string TypeiCode { get; set; }
        public string ReferenceNo { get; set; }
        public string Memo { get; set; }
        public string PostedBy { get; set; }
        public DateTime DatePosted { get; set; }
        public string AuthorizedBy { get; set; }
        public DateTime AuthorizedDate { get; set; }
        public string cpActionType { get; set; }
        public string VoucherNumber { get; set; }
        public string PinCode { get; set; }
        public string HtmlTable1 { get; set; }
        public string StartingRange { get; set; }
        public string EndingRange { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string Attachments { get; set; }

        public List<GLTransactionViewModel> gLTransactionList = new List<GLTransactionViewModel>();
        public List<AddUpdateViewModel> pAddUpdateViewModel = new List<AddUpdateViewModel>();
        public string Payee { get; set; }
        public string Customer { get; set; }
        public string Type { get; set; }
    }
    public class GLTransactionViewModel
    {
        #region GlTransactions
        public int TransactionID { get; set; }
        public int GLReferenceID { get; set; }
        public string Type { get; set; }
        public string AccountTypeiCode { get; set; }
        public string Memo { get; set; }
        public decimal BaseAmount { get; set; }
        public decimal LocalAmount { get; set; }
        public string ForeignCurrencyISOCode { get; set; }
        public decimal ForeignCurrencyAmount { get; set; }
        public decimal LocalCurrencyAmount { get; set; }
        public decimal ExchangeRate { get; set; }
        public string GLPersonTypeiCode { get; set; }
        public string PersonID { get; set; }
        public short DimensionId { get; set; }
        public short Dimension2Id { get; set; }
        public string AddedBy { get; set; }
        public DateTime AddedDate { get; set; }
        public string AccountCode { get; set; }
        public string LocalCurrencyBalance { get; set; }
        #endregion
    }

    public class AccountBalance
    {
        public string AccountCode { get; set; }
        public decimal ForeignCurrencyBalance { get; set; }
        public decimal LocalCurrencyBalance { get; set; }
    }

    public class AccountDetails
    {
        public string AccountCode { get; set; }
        public string LocalCurrencyCode { get; set; }
        public string ForeignCurrencyCode { get; set; }
        public decimal CreditMaxRate { get; set; }
        public decimal CreditMinRate { get; set; }
        public decimal ForeignCurrencyBalance { get; set; }
        public decimal LocalCurrencyBalance { get; set; }
        public decimal CurrencyRate { get; set; }
        public string CurrencyType { get; set; }
    }

    public class AddUpdateViewModel
    {
        public string BankCashAccountCode { get; set; }
        public string GeneralAccountCode { get; set; }
        public decimal LocalCurrencyAmount { get; set; }
        public decimal ForeignCurrencyAmount { get; set; }
        public decimal ExchangeRate { get; set; }
        public string DrCr { get; set; }
        public string cpUserRefNo { get; set; }
        public int cpEntryID { get; set; }
        public decimal CreditMaxRate { get; set; }
        public decimal CreditMinRate { get; set; }
        public string Memo { get; set; }
        public string CurrencyISoCode { get; set; }
    }
    public class ReturnViewModel
    {
        public bool isPosted { get; set; }
        public string ErrMessage { get; set; }
        public string HtmlTable { get; set; }
    }

    public class EditViewModel
    {
        public string cpEntryID { get; set; }
        public string AccountCode { get; set; }
        public decimal ForeignCurrencyAmount { get; set; }
        public decimal LocalCurrencyAmount { get; set; }
        public string Memo { get; set; }
        public decimal CreditMaxRate { get; set; }
        public decimal CreditMinRate { get; set; }
        public string DrCr { get; set; }
        public string ForeignCurrencyISOCode { get; set; }
        public string LocalCurrencyISOCode { get; set; }
        public decimal CurrencyRate { get; set; }
        public string CurrencyType { get; set; }
        public decimal ForeignCurrencyBalance { get; set; }
        public decimal LocalCurrencyBalance { get; set; }


    }
    public class DeleteViewModel
    {
        public bool isPosted { get; set; }
        public string HtmlTable { get; set; }
        public string cpUserRefNo { get; set; }
        public string Attachments { get; set; }

    }
}