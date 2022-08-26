using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class TransViewModel
    {
        public string cpRefID { get; set; }
        public string VoucherNumber { get; set; }
        public string HtmlTable1 { get; set; }
        public string HtmlTable2 { get; set; }
        public string DatePosted { get; set; }
        public string AuthorizedBy { get; set; }
        public string AuthorizedDate { get; set; }
        public string PostedBy { get; set; }
        public string Attachments { get; set; }
        public string TypeiCode { get; set; }
        public string VoucherTitle { get; set; }

    }
}