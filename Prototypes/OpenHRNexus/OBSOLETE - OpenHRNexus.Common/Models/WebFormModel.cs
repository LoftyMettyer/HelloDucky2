using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace OpenHRNexus.Common.Models
{
    public class WebFormModel
    {
        [Key]
        public string form_id { get; set; }
        public string form_name { get; set; }
        public List<WebFormField> form_fields { get; set; }
    }

}
