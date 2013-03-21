using NServiceBus;
using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.MessageSenders
{
    public class StaffContractChangeMessageSender : TrackingMessageSender<StaffContractChangeRequest>
    {

        public IBus Bus { get; set; }

        public override void Send(StaffContractChangeRequest message)
        {
            TrackMessage(message);

            if (!LaterInboundMessageProcessed(message))
            {
                Bus.Send(message);
            }
        }

    }
}


