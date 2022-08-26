using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
namespace Application.DTOs
{
    public class ChartOfAccountsModel
    {
        public int AccountID { get; set; }
        public string AccountCode { get; set; }
        public string AccountName { get; set; }
        public string ParentAccountCode { get; set; }
        public string ParentAccountName { get; set; }
        public string AccountCodeMask { get; set; }
        public bool AccountStatus { get; set; }
        public string AccountDescription { get; set; }
        public string AddedBy { get; set; }
        public DateTime AddedDate { get; set; }
        public int ClassID { get; set; }
        public int TypeID { get; set; }
        //[Required]
        public string Prefix { get; set; }
        //[Required]
        public string CurrencyISOCode { get; set; }
        //[Required]
        public string CountryISONumericCode { get; set; }
        public bool RevalueYN { get; set; }
        public decimal Revalue { get; set; }
        public decimal DebiteLimit { get; set; }
        public decimal CreditLimit { get; set; }
        public decimal TaxPercentage { get; set; }
        public decimal VATWHT { get; set; }
        public bool Charges { get; set; }
        public bool isHead { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public bool isvalid { get; set; }
        public decimal ForeignCurrencyBalance { get; set; }
        public decimal LocalCurrencyBalance { get; set; }
        public string HtmlTable { get; set; }
        public bool ShowAll { get; set; }
        public decimal TotalBalance { get; set; }
        public string tempTable { get; set; }
        public bool isPrint { get; set; }
        public bool isVisible { get; set; }

        public int preAccountID { get; set; }
        public string preAccountCode { get; set; }
        public string preParentAccountCode { get; set; }
        public string preParentAccountName { get; set; }
        public string PinCode { get; set; }
        public string AccountTitle { get; set; }
    }

    public class ChartOfAccountsTree
    {
        public int AccountID { get; set; }
        public int AccountCode { get; set; }
        public int ParentAccountCode { get; set; }
        public string AccountName { get; set; }
        public bool AccountStatus { get; set; }
        public string CurrencyISOCode { get; set; }
        public string Prefix { get; set; }
        public DateTime AddedDate { get; set; }
        public decimal ForeignCurrencyBalance { get; set; }
        public decimal LocalCurrencyBalance { get; set; }

        public bool isHead { get; set; }

    }


    public class BuyerReportViewModel
    {
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string AccountCode { get; set; }
        public string OpeningBalanceDate { get; set; }
        public DateTime TransactionDate { get; set; }
        public int TotalTransactions { get; set; }
        public bool isvalid { get; set; }
        public decimal DebitFCTotal { get; set; }
        public decimal DebitLCTotal { get; set; }
        public decimal CreditFCTotal { get; set; }
        public decimal CreditLCTotal { get; set; }
        public string HtmlTable { get; set; }
        public bool ShowAll { get; set; }
        public decimal TotalFCBalance { get; set; }
        public decimal TotalLCBalance { get; set; }
        public string tempTable { get; set; }
        public int AccountID { get; set; }
        public string AccountName { get; set; }
        public string ErrMessage { get; set; }
        public bool IsPDF { get; set; }
        public DateTime AddedDate { get; set; }
        public string pDate { get; set; }
        public int Total { get; set; }
        public string pTotal { get; set; }
        public decimal LocalCurrencyBalance { get; set; }
        public decimal ForeignCurrencyBalance { get; set; }
        public string VoucherNumber { get; set; }
        public int SortOrder { get; set; }
    }
    public class HoldFileReportViewModel
    {
        public string Date { get; set; }
        public string InvoiceNo { get; set; }
        public string Agent { get; set; }
        public decimal AmountInFC { get; set; }
        public decimal AmountInLC { get; set; }
        public decimal TotalFCBalance { get; set; }
        public decimal TotalLCBalance { get; set; }
        public string tempTable { get; set; }
        public bool isvalid { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string VoucherNo { get; set; }
        public string Memo { get; set; }
        public string TransactionID { get; set; }
        public string Recipient { get; set; }
        public string AgentPrefix { get; set; }
        public string SenderName { get; set; }
        public string CustomerId { get; set; }
        public string Description { get; set; }
        public string PaymentNo { get; set; }
        public string ReceiverName { get; set; }
        public string AccountCode { get; set; }
    }

    public class AgentReportViewModel
    {
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string AccountCode { get; set; }
        public string Country { get; set; }

        public int TotalTransactions { get; set; }
        public bool isvalid { get; set; }
        public decimal AdminFCAmount { get; set; }
        public decimal AdminLCAmount { get; set; }
        public decimal AgentFCAmount { get; set; }
        public decimal AgentLCAmount { get; set; }
        public decimal FXMarginFCAmount { get; set; }
        public decimal FXMarginLCAmount { get; set; }
        public string HtmlTable { get; set; }
        public bool ShowAll { get; set; }
        public decimal TotalFCBalance { get; set; }
        public decimal TotalLCBalance { get; set; }
        public string tempTable { get; set; }
        public int AccountID { get; set; }
        public string AccountName { get; set; }
        public string ErrMessage { get; set; }
        public bool IsPDF { get; set; }
    }
    public class ConfigItemsDataModel
    {
        public int ConfigItemsDataID { get; set; }
        public string ItemsCode { get; set; }
        public string ItemsDataCode { get; set; }
        public string Description { get; set; }
        public bool? Status { get; set; }

    }
    public class RecursiveAccountsViewModel
    {

        public string Path { get; set; }

        public string[] arr { get; set; }

    }
}