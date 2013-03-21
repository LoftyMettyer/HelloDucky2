using Fusion.Connector.OpenHR.MessageSenders;
using NServiceBus;
using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffSkillChange
{
    public class StaffSkillChangeMessageSender : TrackingMessageSender<StaffSkillChangeRequest>
    {
        public IBus Bus { get; set; }

        public override void Send(StaffSkillChangeRequest message)
        {
            TrackMessage(message);

            if (!LaterInboundMessageProcessed(message))
            {
                Bus.Send(message);
            }
        }

    }
}
