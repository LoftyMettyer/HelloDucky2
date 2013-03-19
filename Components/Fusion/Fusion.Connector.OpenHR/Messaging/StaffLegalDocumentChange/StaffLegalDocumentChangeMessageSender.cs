using Fusion.Connector.OpenHR.MessageSenders;
using NServiceBus;
using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffLegalDocumentChange
{
    public class StaffLegalDocumentChangeMessageSender : TrackingMessageSender<StaffLegalDocumentChangeRequest>
    {
        public IBus Bus { get; set; }

        public override void Send(StaffLegalDocumentChangeRequest message)
        {
            TrackMessage(message);

            if (!LaterInboundMessageProcessed(message))
            {
                Bus.Send(message);
            }
        }

    }
}
