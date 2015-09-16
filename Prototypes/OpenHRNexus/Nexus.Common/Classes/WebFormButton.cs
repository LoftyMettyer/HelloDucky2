using Nexus.Common.Enums;

namespace Nexus.Common.Models
{
    public class WebFormButton : WebFormControl
    {
        public string Title { get; set; }
        public string TargetUrl { get; set; }
        public ButtonAction Action { get; set; }
    }
}
