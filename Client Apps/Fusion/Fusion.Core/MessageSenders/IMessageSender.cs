using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;

namespace Fusion.Core.MessageSenders
{

    public interface IMessageSender
    {
        void Send(FusionMessage message);

    }
    public interface IMessageSender<T> where T : FusionMessage
    {
        void Send(T message);
    }
}
