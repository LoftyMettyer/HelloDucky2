// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ServiceUserChangeMessage.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the service user change message class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Messages.SocialCare
{
    using Fusion.Messages.General;
    using NServiceBus;

    public class ServiceUserChangeMessage : FusionMessage, IEvent
    {
    }

    public class ServiceUserChangeRequest : FusionMessage, ICommand
    {
    }

}
