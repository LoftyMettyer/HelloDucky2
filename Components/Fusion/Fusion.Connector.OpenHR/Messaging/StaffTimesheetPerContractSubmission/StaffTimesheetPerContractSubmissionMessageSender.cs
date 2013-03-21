
using Fusion.Connector.OpenHR.MessageSenders;
using NServiceBus;
using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffTimesheetPerContractSubmission
{
    public class StaffTimesheetPerContractSubmissionMessageSender : TrackingMessageSender<StaffTimeSheetPerContractSubmissionMessage>
    {
        public IBus Bus { get; set; }

        public override void Send(StaffTimeSheetPerContractSubmissionMessage message)
        {
            TrackMessage(message);

            if (!LaterInboundMessageProcessed(message))
            {
                Bus.Send(message);
            }
        }

    }
}
