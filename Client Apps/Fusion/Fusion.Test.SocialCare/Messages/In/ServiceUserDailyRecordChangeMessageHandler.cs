// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StaffChangeMessageHandler.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements message handler for appropriate test message
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Test.SocialCare.MessageHandlers
{
    using Fusion.Messages.SocialCare;
    using NServiceBus;
    using Fusion.Core.Test;

    /// <summary>
    /// Handle incoming test message
    /// </summary>
    public class ServiceUserDailyRecordChangeMessageHandler : BaseWriteFileMessageHandler, IHandleMessages<ServiceUserDailyRecordChangeMessage>
    {
        public void Handle(ServiceUserDailyRecordChangeMessage message)
        {
            base.WriteMessage(message);
        }

    }
}
