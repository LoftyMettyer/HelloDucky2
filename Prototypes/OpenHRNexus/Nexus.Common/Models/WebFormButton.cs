using Nexus.Common.Enums;
using System.ComponentModel.DataAnnotations;

namespace Nexus.Common.Models
{
    public class WebFormButton
    {
        [Key]
        public int id { get; set; }
        public string title { get; set; }
        public string targeturl { get; set; }
        public WebForm WebForm { get; set; }

        public ButtonAction action { get; set; }
    }
}
