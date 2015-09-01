﻿using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Nexus.Common.Models
{
    public class WebFormField
    {
        public int id { get; set; }
        public int field_id { get; set; }
        public int field_columnid { get; set; }

     //   Public DynamicColumn Column { get; set; }
        public string field_title { get; set; }
        public string field_type { get; set; }
        public string field_value { get; set; }
        public bool field_required { get; set; }
        public bool field_disabled { get; set; }
        public List<WebFormFieldOption> field_options { get; set; }
        public WebForm WebForm { get; set; }
    }
}
