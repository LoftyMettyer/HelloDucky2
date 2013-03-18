// --------------------------------------------------------------------------------------------------------------------
// <copyright file="PayrollIdAssignedMessageSender.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the payroll identifier assigned message sender class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Connector1.MessageSenders
{
    using Fusion.Core.MessageSenders;
    using Fusion.Messages.Example;
    using NServiceBus;
    using Fusion.Core.Sql;  
    
    /// <summary>
    /// Payroll identifier assigned message sender. 
    /// </summary>
    public class PayrollIdAssignedMessageSender : TrackingMessageSender<PayrollIdAssignedMessage>
    {
        /// <summary>
        /// Gets or sets the bus.
        /// </summary>
        /// <value>
        /// The nservicebus.
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
        public override void Send(PayrollIdAssignedMessage message)
        {
            base.TrackMessage(message);

            this.Bus.Publish(message);
        }
    }
}
