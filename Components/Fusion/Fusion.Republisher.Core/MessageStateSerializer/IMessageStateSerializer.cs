using System;
namespace Fusion.Republisher.Core.MessageStateSerializer
{
    public interface IMessageStateSerializer
    {
        MessagePersistedState Deserialize(string messageState);
        string Serialize(MessagePersistedState state);
    }
}
