using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class GLSettingsViewModel
    {
        public bool isPosted { get; set; }
        public string message { get; set; }
        public string PinCode { get; set; }
        public string DefaultCurrency { get; set; }
        public string AgentBaseAccountCode { get; set; }
        public string BuyerBaseAccountCode { get; set; }
        public string BankBaseAccountCode { get; set; }
        public string BankChargesAccountCode { get; set; }

        public string TranslationGainLossAccountCode { get; set; }
        public string TranslationGainLossAccountName { get; set; }
        public string ReceivableAccountCode { get; set; }
        public string HoldingAccountCode { get; set; }
        public string CommissionAccountCode { get; set; }
        public string ExchangeVariancesAccountCode { get; set; }
        public string UnDepositedAccountCode { get; set; }
        public string InventoryAccountCode { get; set; }
        public string UnearnedServiceFeeAccountCode { get; set; }
        public string AgentBaseAccountName { get; set; }
        public string BuyerBaseAccountName { get; set; }
        public string BankBaseAccountName { get; set; }
        public string BankChargesAccountName { get; set; }
        public string ReceivableAccountName { get; set; }
        public string HoldingAccountName { get; set; }
        public string CommissionAccountName { get; set; }
        public string ExchangeVariancesAccountName { get; set; }
        public string UnDepositedAccountName { get; set; }
        public string InventoryAccountName { get; set; }
        public string UnearnedServiceFeeAccountName { get; set; }
        public string AutoAuthorizedPaidRB { get; set; }
        public string AutoAuthorizedHoldingRB { get; set; }
        public string AutoAuthorizedCancelRB { get; set; }
        public string AutoAuthorizeDepositRB { get; set; }
        public string ImportFileHasPostingDate { get; set; }
        public string AutoTransferAuthorizationBSU { get; set; }
        public string TransactionAdminCommissionAccountCode { get; set; }
        public string TransactionAdminCommissionAccountName { get; set; }
        public string TransactionAgentCommissionAccountCode { get; set; }
        public string TransactionAgentCommissionAccountName { get; set; }

        public string ProfileNLossAdministrativeExpensesAccountName { get; set; }
        public string ProfileNLossAdministrativeExpensesAccountCode { get; set; }

        public string ProfileNLossCostOfSalesAccountName { get; set; }
        public string ProfileNLossCostOfSalesAccountCode { get; set; }

        public string ProfileNLossIncomeAccountName { get; set; }
        public string ProfileNLossIncomeAccountCode { get; set; }

        public string ProfileNLossOtherExpensesAccountName { get; set; }
        public string ProfileNLossOtherExpensesAccountCode { get; set; }

        public string ProfileNLossOtherIncomeAccountName { get; set; }
        public string ProfileNLossOtherIncomeAccountCode { get; set; }
        public DateTime FinancialYearFromDate { get; set; }
        public DateTime FinancialYearToDate { get; set; }
        public string FinancialYearDatePostedBy { get; set; }
        public string UserEmail { get; set; }
        public DateTime PostingStartDate { get; set; }
        public string BSNonCurrentAssetsAccountName { get; set; }
        public string BSNonCurrentAssetsAccountCode { get; set; }
        public string BSCurrentAssetsAccountName { get; set; }
        public string BSCurrentAssetsAccountCode { get; set; }
        public string BSCurrentLiabilitesAccountName { get; set; }
        public string BSCurrentLiabilitesAccountCode { get; set; }
        public string BSNonCurrentliabilitesAccountName { get; set; }
        public string BSNonCurrentliabilitesAccountCode { get; set; }
        public string BSCapitalNReservesAccountName { get; set; }
        public string BSCapitalNReservesAccountCode { get; set; }

    }
    public class AddAccountCodeResponse
    {
        public bool isAdded = false;
        public string AccountCode { get; set; }
        public string AccountName { get; set; }
    }
}