using Fusion.Core.Test;
using Fusion.Messages.General;
using NServiceBus;

namespace Fusion.Test.SocialCare.Messages.In
{
    public class GenericMessageHandler : BaseWriteFileMessageHandler, IHandleMessages<FusionMessage>
    {
        public void Handle(FusionMessage message)
        {
            base.WriteMessage(message);
        }
    }
}
