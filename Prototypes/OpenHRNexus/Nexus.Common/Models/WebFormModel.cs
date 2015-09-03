using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Nexus.Common.Models
{
    public class WebFormModel
    {
        public int id { get; set; }
        [Key]
        public Guid stepid { get; set; }
        public string name { get; set; }
        public List<WebFormField> fields { get; set; }
        public List<WebFormButton> buttons { get; set; }

    }

}
