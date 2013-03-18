// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StaffChangeMessage.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the staff picture change message class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Messages.SocialCare
{
    using Fusion.Messages.General;
    using NServiceBus;

    public class StaffLegalDocumentChangeRequest : FusionMessage, ICommand
    {
    }

    public class StaffLegalDocumentChangeMessage : FusionMessage, IEvent
    {
    }
}
