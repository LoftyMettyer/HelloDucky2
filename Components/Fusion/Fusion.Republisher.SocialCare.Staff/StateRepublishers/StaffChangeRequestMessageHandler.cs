
namespace Fusion.Publisher.SocialCare
{
    using Fusion.Messages.SocialCare;
    using Fusion.Publisher.SocialCare.MessageDefinitions;
    using Fusion.Republisher.Core;

    public class StaffChangeRequestMessageHandler : StateStoreMessageRepublisher<StaffChangeRequest, StaffChangeMessage, StaffChangeMessageDefinition>
    {
        
    }
}
