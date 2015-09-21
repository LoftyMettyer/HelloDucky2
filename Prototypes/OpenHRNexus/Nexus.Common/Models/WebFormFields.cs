using Nexus.Common.Interfaces;
using System.Collections.Generic;

namespace Nexus.Common.Models
{
    public class WebFormField : WebFormControl
    {
        private string _elementId;

        public int sequence { get; set; }

        public int columnid { get; set; }

        /// <summary>
        /// Todo the elementId needs to be properly defined. Temporary while end finalise the endpoint constructs
        /// </summary>
        public string elementid { get; set; }

        public string title { get; set; }
        public string type { get; set; }
        public string value { get; set; }
        public bool required { get; set; }
        public bool disabled { get; set; }
        public List<WebFormFieldOption> options { get; set; }
        public bool IsLookupValue { get; set; }
    }
}
