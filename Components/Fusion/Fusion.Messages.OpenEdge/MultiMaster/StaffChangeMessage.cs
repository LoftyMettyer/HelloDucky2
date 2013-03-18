
namespace Fusion.Messages.OpenEdge
{
    using System;
    using NServiceBus;
    using Fusion.Messages.General;
    
    public class StaffChangeRequest : FusionMessage, ICommand
    {
    }
    
    public class StaffChangeMessage : FusionMessage, IEvent
    {
    }
}
