using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Core.MessageSenders;
using Fusion.Messages.General;
using Fusion.Core.Sql;
using StructureMap.Attributes;
using Fusion.Core;

namespace Connector1.MessageSenders
{
    public abstract class TrackingMessageSender<T> : MessageSender<T> where T: FusionMessage
    {
        [SetterProperty]
        public IMessageTracking MessageTracking {
        
            get;
            set;

        }

        protected void TrackMessage(T message)
        {
            MessageTracking.SetLastGeneratedDate(message.GetMessageName(), message.EntityRef.Value, message.CreatedUtc);
            MessageTracking.SetLastGeneratedXml(message.GetMessageName(), message.EntityRef.Value, message.Xml);
        }


    }
}
