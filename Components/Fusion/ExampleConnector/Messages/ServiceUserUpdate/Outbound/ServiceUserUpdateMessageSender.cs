// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ServiceUserUpdateMessageSender.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the service user update message sender class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Connector1.MessageSenders
{
    using Fusion.Core.MessageSenders;
    using Fusion.Messages.Example;
    using NServiceBus;
    using Fusion.Core.Sql;

    /// <summary>
    /// Service user update message sender. 
    /// </summary>
    public class ServiceUserUpdateMessageSender : TrackingMessageSender<ServiceUserUpdateRequest>
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
        public override void Send(ServiceUserUpdateRequest message)
        {
            base.TrackMessage(message);

            if (!base.LaterInboundMessageProcessed(message))
            {
                this.Bus.Send(message);
            }
        }
    }
}
