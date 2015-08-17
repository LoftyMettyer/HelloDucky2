using System.Collections.Generic;

namespace OpenHRNexus.Common.Models
{
    public class WebFormModel
    {
        public string form_id { get; set; }
        public string form_name { get; set; }
        public List<WebFormFields> form_fields { get; set; }
    }

}
