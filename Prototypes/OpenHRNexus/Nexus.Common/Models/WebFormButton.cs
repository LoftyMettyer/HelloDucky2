using System.ComponentModel.DataAnnotations;

namespace Nexus.Common.Models
{
    public class WebFormButton
    {
        [Key]
        public int button_id { get; set; }
        public string button_title { get; set; }
        public string button_targeturl { get; set; }
        public WebForm WebForm { get; set; }
    }
}
