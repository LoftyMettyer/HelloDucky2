using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Core.Sql;
using StructureMap.Attributes;
using Fusion.Core.OutboundFilters;
using Fusion.Messages.General;
using Fusion.Core;

namespace Connector1.OutboundFilters
{
    public abstract class CheckDuplicateOutboundFilterHandler<T> : OutboundFilterHandler<T> where T: FusionMessage
    {
        
        [SetterProperty]
        public IMessageTracking MessageTracking {
            get;
            set;
        }

        public override bool Handle(T message)
        {
            string messageType = message.GetMessageName();

            Guid busRef = message.EntityRef.Value;              

            string lastMessageGenerated = this.MessageTracking.GetLastGeneratedXml(messageType, busRef);

            // Simple compare, could be more elaborate
            if (lastMessageGenerated == message.Xml)
                return false;
            
            return true;
        }

    }
}
