using System.Collections.Generic;

namespace OpenHRNexus.Common.Models
{
    public class WebForm
    {
        public int id { get; set; }
        public string Name { get; set; }
        public List<WebFormField> Fields { get; set; }

    }

}
