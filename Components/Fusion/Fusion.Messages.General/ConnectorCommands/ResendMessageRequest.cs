

namespace Fusion.Messages.General.ConnectorCommands
{
    using NServiceBus;
    using System;

    /// <summary>
    /// Command to send out a particular message (to standard destinations)
    /// </summary>
    public class ResendMessageRequest : ICommand
    {
        public string Community;
        public string MessageType;
        public Guid EntityRef;
    }
}
