using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class PostingDateViewModel
    {
        public bool isPosted { get; set; }
        public DateTime MaxDate { get; set; }
        public string UserMenuID { get; set; }
        public DateTime PostingDate { get; set; }
        public string message { get; set; }
        public string PinCode { get; set; }
        //public DateTime MaxDate { get; set; }


    }
}