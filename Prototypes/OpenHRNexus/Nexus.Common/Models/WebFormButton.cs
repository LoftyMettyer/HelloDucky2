using Nexus.Common.Enums;

namespace Nexus.Common.Models
{
    public class WebFormButton : WebFormControl
    {
        public string title { get; set; }
        public string targeturl { get; set; }
        public ButtonAction action { get; set; }
    }
}
