using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Nexus.Common.Models
{
    public class WebFormModel
    {
        [Key]
        public string form_id { get; set; }
        public string form_name { get; set; }
        public List<WebFormField> form_fields { get; set; }
        public List<WebFormButton> form_buttons { get; set; }

    }

}
