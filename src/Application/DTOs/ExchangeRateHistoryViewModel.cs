﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class ExchangeRateHistoryViewModel
    {
        public string CurrencyISOCode { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string Table { get; set; }
        public string ErrMessage { get; set; }

    }
}