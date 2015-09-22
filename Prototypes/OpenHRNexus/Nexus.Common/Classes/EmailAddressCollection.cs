using System.Net.Mail;

namespace Nexus.Common.Classes
{
    public class EmailAddressCollection
    {
        public string To { get; set; }
        public string CC { get; set; }
        public string BCC { get; set; }
        public string ReplyTo { get; set; }
        public string From { get; set; }

    }
}
