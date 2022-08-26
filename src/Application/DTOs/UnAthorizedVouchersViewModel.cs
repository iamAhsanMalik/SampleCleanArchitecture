using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class UnAuthorizedVouchersViewModel
    {
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string ErrMessage { get; set; }
        public string Table { get; set; }
        public string UnAuthorizedAllVouchers { get; set; }
        public string UnAuthorizedCashVouchers { get; set; }
        public string UnAuthorizedTransferVouchers { get; set; }




    }
}