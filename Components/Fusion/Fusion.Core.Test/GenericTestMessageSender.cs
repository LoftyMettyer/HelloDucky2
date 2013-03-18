using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Core.MessageSenders;
using Fusion.Messages.General;
using NServiceBus;

namespace Fusion.Core.Test
{
    public class GenericTestMessageSender : IMessageSender
    {
        /// <summary>
        /// Gets or sets the bus (injected)
        /// </summary>
        /// <value>
        /// The message bus
        /// </value>
        public IBus Bus
        {
            get;
            set;
        }

        /// <summary>
        /// Send this message.
        /// </summary>
        /// <param name="message"> The message. </param>
        public void Send(FusionMessage message)
        {
            if (message is ICommand)
            {
                this.Bus.Send(message);
            }
            else if (message is IEvent)
            {
                this.Bus.Publish(message);
            }
            else
            {
                throw new InvalidOperationException(String.Format("I don't know how to send a message of type {0}", message.GetType()));
            }
        }
    }
}
