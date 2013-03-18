using Fusion.Republisher.Core.MessageStateSerializer;
using System;

namespace Fusion.Republisher.Core.MessageProcessors
{
    public interface IFusionMessageProcessor
    {
        string CreateMessageFromState(IMessageDefinition messageDefinition, MessagePersistedState currentState);
        void UpdateStateFromMessage(IMessageDefinition messageDefinition, MessagePersistedState currentState, string messageXml);
    }
}
