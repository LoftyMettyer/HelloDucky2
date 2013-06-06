// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StaffChangeRequestMessageSender.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the class that will send messages of this type to the bus
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Test.SocialCare.MessageSenders
{
    using Fusion.Core.MessageSenders;
    using Fusion.Messages.SocialCare;
    using NServiceBus;

    /// <summary>
    /// Send messages of this type to the bus. 
    /// </summary>
    public class StaffPostChangeRequestMessageSender : MessageSender<StaffPostChangeRequest>
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
        public override void Send(StaffPostChangeRequest message)
        {
            this.Bus.Send(message);
        }
    }
}
