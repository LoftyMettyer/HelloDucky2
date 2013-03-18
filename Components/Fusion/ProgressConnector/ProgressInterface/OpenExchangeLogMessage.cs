using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProgressConnector.ProgressInterface
{
    public class OpenExchangeLogMessage : OpenExchangeGeneratedContent
    {
        public Guid Id { get; set; }

        public string Source { get; set; }

        public Guid? MessageId { get; set; }

        public Guid? EntityRef { get; set; }

        public Guid? PrimaryEntityRef { get; set; }
        public DateTime TimeUtc { get; set; }

        public string LogLevel { get; set; }

        public string Message { get; set; }

        public string MessageDescription { get; set; }

        public string Community
        {
            get;
            set;
        }
    }
}
