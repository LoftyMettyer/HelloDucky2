using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebAPI3.Models
{
    public class AbsenceRequest :Absence
    {
        public DateTime? RequestDate { get; set; }
    }
}