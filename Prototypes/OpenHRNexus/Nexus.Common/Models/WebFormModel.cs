using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Nexus.Common.Models
{
    public class WebFormModel
    {
        [Key]
        public string id { get; set; }
        public string name { get; set; }
        public List<WebFormField> fields { get; set; }
        public List<WebFormButton> buttons { get; set; }

    }

}
