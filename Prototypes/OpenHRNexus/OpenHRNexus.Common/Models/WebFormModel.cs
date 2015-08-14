using System.Collections.Generic;

namespace OpenHRNexus.Common.Models
{
    public class WebFormModel
    {
        public string form_id { get; set; }
        public string form_name { get; set; }
        public List<WebFormFields> form_fields { get; set; }
    }

    public class WebFormFields
    {
        public int field_id { get; set; }
        public string field_title { get; set; }
        public string field_type { get; set; }
        public string field_value { get; set; }
        public bool field_required { get; set; }
        public bool field_disabled { get; set; }
        public List<WebFormFieldOption> field_options { get; set; }

    }

    public class WebFormFieldOption
    {
        public int option_id { get; set; }
        public string option_title { get; set; }
        public int option_value { get; set; }
    }
}
