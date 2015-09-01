using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Nexus.Common.Models
{
    public class WebFormField
    {
        public int id { get; set; }
        public int sequence { get; set; }
        public int columnid { get; set; }

     //   Public DynamicColumn Column { get; set; }
        public string title { get; set; }
        public string type { get; set; }
        public string value { get; set; }
        public bool required { get; set; }
        public bool disabled { get; set; }
        public List<WebFormFieldOption> options { get; set; }
        public WebForm WebForm { get; set; }
    }
}
