

namespace Fusion.Republisher.SocialCare
{
    using Fusion.Messages.SocialCare;
    using log4net;
    using NServiceBus;

    public class StaffPostChangeRequestMessageHandler : IHandleMessages<StaffPostChangeRequest>
    {

        public IBus Bus
        {
            get;
            set;
        }

        public void Handle(StaffPostChangeRequest message)
        {
            Logger.Info(string.Format("Fusion Publisher received StaffPostChangeRequest with Id {0}.", message.Id));
            Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));

            Logger.Info(string.Format("Message name: {0}.", message.GetType()));
            Logger.Info(string.Format("Message source: {0}.", message.Originator));

            Logger.Info("Republishing to bus");

            // Blind republish
            Bus.Publish(new StaffPostChangeMessage
            {
                CreatedUtc = message.CreatedUtc,
                EntityRef = message.EntityRef,
                Id = message.Id,
                Originator = message.Originator,
                SchemaVersion = message.SchemaVersion,
                Xml = message.Xml
            }
            );

            //Logger.Info("Persisting to local database");
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(StaffPostChangeRequestMessageHandler));
    }


}
