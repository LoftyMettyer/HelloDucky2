using NServiceBus;
using Fusion.Messages.General;
using Fusion.Messages.SocialCare;
using StructureMap.Attributes;

namespace Fusion.Connector.OpenHR.MessageSenders
{
    public class StaffChangeMessageSender : TrackingMessageSender<StaffChangeRequest>
    {
        
        public IBus Bus {get;set;}

        public override void Send(StaffChangeRequest message)
        {
            base.TrackMessage(message);

            if (!base.LaterInboundMessageProcessed(message))
            {
                this.Bus.Send(message);
            }
        }

    }
}
