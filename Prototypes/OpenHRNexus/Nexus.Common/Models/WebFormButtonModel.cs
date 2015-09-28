using Nexus.Common.Enums;
using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace Nexus.Common.Models
{
    public class WebFormButtonModel
    {
        public int id { get; set; }
        public string title { get; set; }
        public string TargetUrl { get; set; }
        public ButtonAction action { get; set; }
//        public string Html { get; set; }
        [NotMapped]
//        public string AuthenticationCode { get; set; }
        public Guid TargetStep { get; set; }
    }
}
