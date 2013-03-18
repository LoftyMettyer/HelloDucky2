using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProgressConnector.ProgressInterface
{
    public class OpenExchangeIdTranslation : OpenExchangeGeneratedContent
    {
        public Guid Id { get; set; }

        public string Source { get; set; }

        public Guid? MessageId { get; set; }

        public DateTime TimeUtc { get; set; }

        public string EntityName { get; set; }

        public Guid BusRef { get; set; }

        public string LocalId { get; set; }

        public string Community { get; set; }
    }
}
