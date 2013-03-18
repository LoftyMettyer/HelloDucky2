// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SubscribeToPublications.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the subscribe to publications class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Connector1
{
    using Fusion.Messages.Example;
    using NServiceBus;

    /// <summary>
    /// Showing how to manage subscriptions manually
    /// </summary>
    public class SubscribeToPublications : IWantToRunAtStartup
    {
        /// <summary>
        /// Gets or sets the bus.
        /// </summary>
        /// <value>
        /// The NServiceBus bus
        /// </value>
        public IBus Bus { get; set; }

        /// <summary>
        /// Runs this object.
        /// </summary>
        public void Run()
        {
            this.Bus.Subscribe<ServiceUserUpdateMessage>();
        }

        /// <summary>
        /// Stops this object.
        /// </summary>
        public void Stop()
        {
            this.Bus.Unsubscribe<ServiceUserUpdateMessage>();
        }
    }
}
