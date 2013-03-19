//using Fusion.Messages.General;
//using NServiceBus;
//using Fusion.Core.Sql;
//using Fusion.Messages.SocialCare;

//namespace Fusion.Connector.OpenHR.Messaging
//{
//    public class StaffChangeRequestSender : TrackingMessageSender<StaffChangeRequest>
//    {

//        public IBus Bus { get; set; }

//        public override void Send(StaffChangeRequest message)
//        {
//            base.TrackMessage(message);

//            if (!base.LaterInboundMessageProcessed(message))
//            {

//                if (message is ICommand)
//                {
//                    this.Bus.Send(message);
//                }
//                if (message is IEvent)
//                {
//                    this.Bus.Publish(message);
//                }
//            }
//        }
//    }
//}