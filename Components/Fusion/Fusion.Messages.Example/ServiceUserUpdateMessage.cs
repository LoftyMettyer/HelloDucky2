namespace Fusion.Messages.Example
{
    using System;
    using Fusion.Messages.General;
    using NServiceBus;


    // Multi-master, so just a fusion message

    public class ServiceUserUpdateRequest : FusionMessage, ICommand
    {

    }

    public class ServiceUserUpdateMessage : FusionMessage, IEvent
    {

    }

}