using System.ComponentModel.DataAnnotations;

namespace OpenHRNexus.Common.Models
{
    public class WebFormFieldOption
    {
        [Key]
        public int option_id { get; set; }
        public string option_title { get; set; }
        public int option_value { get; set; }
    }
}
