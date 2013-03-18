

namespace Fusion.Republisher.SocialCare
{
    using Fusion.Messages.SocialCare;
    using Fusion.Publisher.SocialCare.MessageDefinitions;
    using Fusion.Republisher.Core;

    public class ServiceUserChangeRequestMessageHandler : StateStoreMessageRepublisher<ServiceUserChangeRequest, ServiceUserChangeMessage, ServiceUserChangeMessageDefinition>
    {
    
    }
    
}
