using Nexus.Common.Interfaces;
using System.Collections.Generic;

namespace Nexus.Common.Models
{
    public class WebFormField : WebFormControl
    {

        public int sequence { get; set; }
        public int columnid { get; set; }
        public string elementid { get
            {
                return string.Format("we_{0}_{1}", id.ToString(), sequence.ToString());
            }
        }
        public string title { get; set; }
        public string type { get; set; }
        public string value { get; set; }
        public bool required { get; set; }
        public bool disabled { get; set; }
        public List<WebFormFieldOption> options { get; set; }
        public bool IsLookupValue { get; set; }
    }
}
