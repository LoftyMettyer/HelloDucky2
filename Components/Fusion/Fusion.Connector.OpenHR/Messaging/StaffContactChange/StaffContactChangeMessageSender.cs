using NServiceBus;
using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.MessageSenders
{
    public class StaffContactChangeMessageSender: TrackingMessageSender<StaffContactChangeRequest>
    {

        public IBus Bus {get;set;}

        public override void Send(StaffContactChangeRequest message)
        {
            TrackMessage(message);

            if (!LaterInboundMessageProcessed(message))
            {
                Bus.Send(message);
            }
        }

    }
}


