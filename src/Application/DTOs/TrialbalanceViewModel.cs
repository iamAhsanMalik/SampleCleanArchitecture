using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class TrialbalanceViewModel
    {
        public string Table { get; set; }
        public bool ShowAll { get; set; }
        public bool haveData { get; set; }
        public DateTime DateFrom { get; set; }
        public DateTime DateTo { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string AccountName { get; set; }
        public string subName3 { get; set; }
        public string subName2 { get; set; }
        public string subName1 { get; set; }
        public string AccountHead { get; set; }
        public string Debit { get; set; }
        public string Credit { get; set; }
        public string AccountCode { get; set; }
        public string LocalCurrencyBalance { get; set; }
    }
}