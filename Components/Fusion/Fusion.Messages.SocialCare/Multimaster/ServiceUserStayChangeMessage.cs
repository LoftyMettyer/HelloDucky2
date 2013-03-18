// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ServiceUserChangeMessage.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the service user stay change message class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Messages.SocialCare
{
    using Fusion.Messages.General;
    using NServiceBus;

    public class ServiceUserStayChangeMessage : FusionMessage, IEvent
    {
    }

    public class ServiceUserStayChangeRequest : FusionMessage, ICommand
    {
    }

}
