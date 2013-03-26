using NServiceBus;
using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.MessageSenders
{
    public class StaffChangeMessageSender : TrackingMessageSender<StaffChangeRequest>
    {
        
        public IBus Bus {get;set;}

        public override void Send(StaffChangeRequest message)
        {
            TrackMessage(message);

            if (LaterInboundMessageProcessed(message)) return;
            Bus.Send(message);
        }
    }
}
