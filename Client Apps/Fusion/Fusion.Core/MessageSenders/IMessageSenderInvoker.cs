using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;

namespace Fusion.Core.MessageSenders
{
    public interface IMessageSenderInvoker
    {
        void Invoke(FusionMessage message);
    }
}
