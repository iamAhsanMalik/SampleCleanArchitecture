using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class InquiriesViewModel
    {
        public string AccountCode { get; set; }
        public string ExcelAccountCode { get; set; }
        public string PdfAccountCode { get; set; }
        public string GLReferenceID { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public DateTime ExcelFromDate { get; set; }
        public DateTime ExcelToDate { get; set; }
        public DateTime PdfFromDate { get; set; }
        public DateTime PdfToDate { get; set; }

        public string Dimension { get; set; }
        public string CurrencyType { get; set; }
        public string ExcelCurrencyType { get; set; }
        public string PdfCurrencyType { get; set; }
        public decimal MinAmount { get; set; }
        public decimal MaxAmount { get; set; }
        public string table { get; set; }
        public string ErrMessage { get; set; }
        public bool haveData { get; set; }
        public string AccountCodeWithMemo { get => $"{AccountCode} - {Memo}"; }
        public string VoucherNumber { get; set; }
        public string Memo { get; set; }
        public DateTime AddedDate { get; set; }
        public string ForeignCurrencyISOCode { get; set; }
        public string LocalCurrencyISOCode { get; set; }
        public string TransactionNumber { get; set; }
        public string TransactionID { get; set; }

        public decimal PayoutAmount { get; set; }
        public decimal DebitForeignCurrencyAmount { get; set; }
        public decimal DebitLocalCurrencyAmount { get; set; }
        public decimal CreditForeignCurrencyAmount { get; set; }
        public decimal CreditLocalCurrencyAmount { get; set; }
        public decimal LocalCurrencyBalance { get; set; }
        public decimal ForeignCurrencyBalance { get; set; }

        public decimal LocalCurrencyAmount { get; set; }
        public decimal ForeignCurrencyAmount { get; set; }

        public int SortOrder { get; set; }

    }
    public class JournalInquiryViewModel
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
        public string Memo { get; set; }
        public string VoucherNumber { get; set; }
        public string AccountTypeICode { get; set; }
        public string Table { get; set; }
        public string ErrMessage { get; set; }
        public bool haveData { get; set; }
        public string AddedDate { get; set; }
        public string TypeiCode { get; set; }
        public string AuthBy { get; set; }
        public string PostedBy { get; set; }
        public decimal LocalCurrencyAmount { get; set; }
        public decimal ForeignCurrencyAmount { get; set; }
        public string ForeignCurrencyISOCode { get; set; }
        public string AccountTitle { get; set; }
        public string GLReferenceID { get; set; }
        public string AccountCode { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
    }
    public class TrailBalInquiryViewModel
    {
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
    }
}