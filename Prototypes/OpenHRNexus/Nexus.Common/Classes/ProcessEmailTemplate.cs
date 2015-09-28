using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Common.Classes
{
    public class ProcessEmailTemplate
    {
        public int Id { get; set; }
        public EmailAddressCollection Destinations { get; set; }
        public string Body { get; set; }
        public string Subject { get; set; }
        public List<WebFormButtonModel> FollowOnActions { get; set; }
        //public string Region { get; set; }
        //public string Language { get; set; }
    }
}
