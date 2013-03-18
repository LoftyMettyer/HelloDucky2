// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SubscriptionAuthorizer.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the subscription authorizer class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace ExampleConnector
{
    using System.Collections.Generic;
    using NServiceBus;


    /// <summary>
    /// Subscription authorizer. 
    /// </summary>
    public class SubscriptionAuthorizer : IAuthorizeSubscriptions
    {

        /// <summary>
        /// Authorize subscribe.
        /// </summary>
        /// <param name="messageType">    Type of the message. </param>
        /// <param name="clientEndpoint"> The client endpoint. </param>
        /// <param name="headers">        The headers. </param>
        /// <returns>
        /// Authorise all new inbound subscription requests
        /// </returns>
        public bool AuthorizeSubscribe(string messageType, string clientEndpoint, IDictionary<string, string> headers)
        {
            return true;
        }

        /// <summary>
        /// Authorize unsubscribe.
        /// </summary>
        /// <param name="messageType">    Type of the message. </param>
        /// <param name="clientEndpoint"> The client endpoint. </param>
        /// <param name="headers">        The headers. </param>
        /// <returns>
        /// Authorise all unsubscribe requests
        /// </returns>
        public bool AuthorizeUnsubscribe(string messageType, string clientEndpoint, IDictionary<string, string> headers)
        {
            return true;
        }
    }
}
