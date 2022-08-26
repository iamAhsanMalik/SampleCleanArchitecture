using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class CurrencyModel
    {
        public string CurrencyISOCode { get; set; }
        public string CountryISOCode { get; set; }
        public string CurrencyName { get; set; }
        public string CurrencySymbol { get; set; }
        public bool CurrencyStatus { get; set; }
        public DateTime DateAdded { get; set; }
        public string AddedBy { get; set; }
        public string HoldAccountCode { get; set; }
        public string HoldAccountName { get; set; }
        public decimal MinRateLimit { get; set; }
        public decimal MaxRateLimit { get; set; }
        public string CurrenyType { get; set; }
        public DateTime LastUpdatedDate { get; set; }
        public decimal CurrencyRate { get; set; }
        public bool LockYesNo { get; set; }
        public string UnEarnedAgentAccountCode { get; set; }
        public string UnEarnedAdminAccountCode { get; set; }
        public bool isProfit { get; set; }
        public bool RevalueYN { get; set; }
        public string LastUpdatedBy { get; set; }
        public bool CreateHoldAccount { get; set; }
        public bool CreateAdminFeeAccount { get; set; }
        public bool CreateAgentFeeAccount { get; set; }
        public string UnEarnedAgentAccountName { get; set; }
        public string UnEarnedAdminAccountName { get; set; }

        public decimal RevalueExchangeRate { get; set; }
    }
}