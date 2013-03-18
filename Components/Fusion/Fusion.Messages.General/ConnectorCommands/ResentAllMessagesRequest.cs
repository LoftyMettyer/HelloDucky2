
namespace Fusion.Messages.General.ConnectorCommands
{
    using NServiceBus;

    /// <summary>
    /// Command to send out all messages of a particular type
    /// </summary>
    public class ResentAllMessagesRequest : ICommand
    {
        public string Community;
        public string MessageType;
    }
}
