using Fusion.Connector.OpenHR.MessageSenders;
using NServiceBus;
using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffPictureChange
{
    public class StaffPictureChangeMessageSender : TrackingMessageSender<StaffPictureChangeRequest>
    {
        public IBus Bus { get; set; }

        public override void Send(StaffPictureChangeRequest message)
        {
            TrackMessage(message);

            if (!LaterInboundMessageProcessed(message))
            {
                Bus.Send(message);
            }
        }

    }
}
