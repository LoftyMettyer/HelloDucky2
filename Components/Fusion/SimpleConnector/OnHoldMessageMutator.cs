using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NServiceBus.MessageMutator;
using NServiceBus;
using Fusion.Messages.General;

namespace Subscriber1
{
    [Serializable]
    public class OnHoldMessage : IMessage
    {
        public FusionMessage Message
        {
            get;
            set;
        }
    }

    public class OnHoldMessageMutator : IMutateIncomingMessages
    {
        public object MutateIncoming(object message)
        {
            OnHoldMessage oh = message as OnHoldMessage;
            if (oh == null) return message;

            return oh.Message;
        }
    }
}
