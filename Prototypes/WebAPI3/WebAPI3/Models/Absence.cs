using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebAPI3.Models
{
    public class Absence
    {
        public int Id { get; set; }
        public DateTime? Start_Date {get; set;}
        public DateTime? End_Date { get; set; }
        public string Type { get; set; }

    }
}