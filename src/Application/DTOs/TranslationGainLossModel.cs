using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class TranslationGainLossModel
    {
        public string Table { get; set; }
        public string PinCode { get; set; }
        public string cpUserRefNo { get; set; }
        public string Memo { get; set; }
        public string ErrMessage { get; set; }
        public bool isPosted { get; set; }
        public string cpRefID { get; set; }

        public bool haveData { get; set; }


    }
}