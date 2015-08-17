using System.Collections.Generic;

namespace OpenHRNexus.Common.Models
{
    public class WebFormFields
    {
        public int field_id { get; set; }
        public int field_columnid { get; set; }
        public string field_title { get; set; }
        public string field_type { get; set; }
        public string field_value { get; set; }
        public bool field_required { get; set; }
        public bool field_disabled { get; set; }
        public List<WebFormFieldOption> field_options { get; set; }

    }
}
