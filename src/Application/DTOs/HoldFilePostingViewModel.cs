using Kendo.Mvc.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class HoldFilePostingViewModel
    {
        public bool isError { get; set; }
        public bool isFileImported { get; set; }
        public bool isRowValid { get; set; }
        public string RowErrMessage { get; set; }
        public bool isPosted { get; set; }
        public string cpUserRefNo { get; set; }
        public int Id { get; set; }
        public string Sheet { get; set; }
        public string ErrMessage { get; set; }
        public bool HasHeader { get; set; }
        public string ExcelFilePath { get; set; }
        public int RowNumber { get; set; }
        public DateTime TransactionDate { get; set; }
        public string AgentName { get; set; }
        public decimal Rate { get; set; }
        public string PayoutCurrency { get; set; }
        public decimal FCAmount { get; set; }
        public decimal Payinamount { get; set; }
        public decimal AdminCharges { get; set; }
        public decimal AgentCharges { get; set; }
        public int TRANID { get; set; }
        public int CustomerId { get; set; }
        public string CustomerName { get; set; }
        public string FatherName { get; set; }
        public string Recipient { get; set; }
        public string Phone { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Code { get; set; }
        public string PaymentNo { get; set; }
        public string BenePaymentMethod { get; set; }
        public string SendingCountry { get; set; }
        public string ReceivingCountry { get; set; }
        public string Status { get; set; }
        public DateTime PostingDate { get; set; }
        //public List<Grid> grid = new List<Grid>();

        public List<Sheets> sheets = new List<Sheets>();
    }
    public class Sheets
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }
    public class ImportDepositFileViewModel
    {
        public bool isError { get; set; }
        public bool isFileImported { get; set; }
        public string ErrorMessage { get; set; }
        public bool isRowValid { get; set; }
        public string RowErrMessage { get; set; }
        public bool isPosted { get; set; }
        public string cpUserRefNo { get; set; }
        public int Id { get; set; }
        public string Sheet { get; set; }
        public string ErrMessage { get; set; }
        public string AgentPrefix { get; set; }
        public int CustomerId { get; set; }
        public string CustomerName { get; set; }
        public string Currency { get; set; }
        public decimal Amount { get; set; }
        public decimal Rate { get; set; }
        public string Type { get; set; }
        public string Details { get; set; }
        public string Account { get; set; }
        public decimal FCAmount { get; set; }
        public decimal LCAmount { get; set; }
        public string ExcelFilePath { get; set; }
        public int RowNumber { get; set; }
        public string Name { get; set; }
        public string Extension { get; set; }
        public long Size { get; set; }
        public DateTime PostingDate { get; set; }
        public List<ImportDepositFileViewModel> uploadFilesList { get; set; }

    }

    public class ImportPaidFileViewModel
    {
        public bool isError { get; set; }
        public bool isFileImported { get; set; }
        public bool isRowValid { get; set; }
        public string RowErrMessage { get; set; }
        public bool isPosted { get; set; }
        public string cpUserRefNo { get; set; }
        public int Id { get; set; }
        public string Sheet { get; set; }
        public string ErrMessage { get; set; }
        public bool HasHeader { get; set; }
        public string ExcelFilePath { get; set; }
        public int RowNumber { get; set; }
        public DateTime TransactionDate { get; set; }
        public DateTime PostingDate { get; set; }
        public string AgentName { get; set; }
        public decimal Rate { get; set; }
        public string PayoutCurrency { get; set; }
        public decimal FCAmount { get; set; }
        public decimal Payinamount { get; set; }
        public decimal AdminCharges { get; set; }
        public decimal AgentCharges { get; set; }
        public int TRANID { get; set; }
        public int CustomerId { get; set; }
        public string CustomerName { get; set; }
        public string FatherName { get; set; }
        public string Recipient { get; set; }
        public string Phone { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Code { get; set; }
        public string PaymentNo { get; set; }
        public string BenePaymentMethod { get; set; }
        public string SendingCountry { get; set; }
        public string ReceivingCountry { get; set; }
        public string Status { get; set; }
        public string Buyer { get; set; }
        public decimal BuyerRate { get; set; }
        public decimal BuyerRateSC { get; set; }
        public decimal BuyerRateDC { get; set; }

        public List<Sheets> sheets = new List<Sheets>();
    }
    public class ImportCancelledFileViewModel
    {
        public bool isError { get; set; }
        public bool isFileImported { get; set; }
        public bool isRowValid { get; set; }
        public string RowErrMessage { get; set; }
        public bool isPosted { get; set; }
        public string cpUserRefNo { get; set; }
        public int Id { get; set; }
        public string Sheet { get; set; }
        public string ErrMessage { get; set; }
        public bool HasHeader { get; set; }
        public string ExcelFilePath { get; set; }
        public int RowNumber { get; set; }
        public DateTime TransactionDate { get; set; }
        public string AgentName { get; set; }
        public decimal Rate { get; set; }
        public string PayoutCurrency { get; set; }
        public decimal FCAmount { get; set; }
        public decimal Payinamount { get; set; }
        public decimal AdminCharges { get; set; }
        public decimal AgentCharges { get; set; }
        public decimal Charges { get; set; }
        public int TRANID { get; set; }
        public int CustomerId { get; set; }
        public string CustomerName { get; set; }
        public string FatherName { get; set; }
        public string Recipient { get; set; }
        public string Phone { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Code { get; set; }
        public string PaymentNo { get; set; }
        public string BenePaymentMethod { get; set; }
        public string SendingCountry { get; set; }
        public string ReceivingCountry { get; set; }
        public string Status { get; set; }
        public DateTime PostingDate { get; set; }

        //public List<Grid> grid = new List<Grid>();

        public List<Sheets> sheets = new List<Sheets>();
    }
}