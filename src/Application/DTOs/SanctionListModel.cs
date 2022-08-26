using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Application.DTOs
{
    public class SanctionListModel
    {
        public string FileName { get; set; }
        public string FileType { get; set; }
        public string AddedDate { get; set; }
        public string AddedBy { get; set; }
        public string FileUrl { get; set; }
    }
}