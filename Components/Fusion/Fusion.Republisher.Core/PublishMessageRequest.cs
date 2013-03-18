using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Republisher.Core
{
    public class PublishMessageRequest
    {
        public Guid EntityRef
        {
            get;
            set;
        }

        public string Originator
        {
            get;
            set;
        }

        public DateTime TriggerDateUtc
        {
            get;
            set;
        }
    }
}
