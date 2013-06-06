// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IServiceBrokerListener.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Declares the IServiceBrokerListener interface
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Core.Sql.ServiceBroker
{
    using System;

    /// <summary>
    /// Interface for service broker listener. 
    /// </summary>
    public interface IFusionServiceBrokerListener
    {
        /// <summary>
        /// Received a single send fusion message-request message from service broker (null for timeout)
        /// </summary>
        /// <returns>
        /// Message request
        /// </returns>
        SendFusionMessageRequest ReceiveMessage();
    }
}
