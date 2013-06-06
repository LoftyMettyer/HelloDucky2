using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;
using StructureMap;

namespace Fusion.Core.MessageSenders
{
    public class MessageSenderInvoker : IMessageSenderInvoker
    {
        public void Invoke(FusionMessage message)
        {
            var messageType = message.GetType();
            var handlerType = typeof(IMessageSender<>).MakeGenericType(messageType);

            var genericHandlerMatches = ObjectFactory.GetAllInstances(handlerType).Cast<IMessageSender>();

            if (genericHandlerMatches.Count() != 1)
            {
                throw new ArgumentException("Cannot locate unique IMessageSender<T> for type", "message");
            }

            IMessageSender sender = genericHandlerMatches.First();

            sender.Send(message);
        }
    }
}
