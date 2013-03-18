﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StaffChangeMessage.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the staff change message class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Messages.SocialCare
{
    using Fusion.Messages.General;
    using NServiceBus;

    public class StaffContactChangeRequest : FusionMessage, ICommand
    {
    }

    public class StaffContactChangeMessage : FusionMessage, IEvent
    {
    }
}
