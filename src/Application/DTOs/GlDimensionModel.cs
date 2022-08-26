using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class GlDimensionModel
    {
        public int DimensionID { get; set; }
        public string Title { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? RequiredDate { get; set; }
        public string Notes { get; set; }
        public bool? Status { get; set; }
    }
}