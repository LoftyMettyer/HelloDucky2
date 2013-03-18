using log4net;

using NServiceBus;
using Fusion.Messages.Example;

namespace Subscriber1
{
    public class ServiceUserUpdateMessageHandler : IHandleMessages<ServiceUserUpdateMessage>
    {
        public void Handle(ServiceUserUpdateMessage message)
        {
            Logger.Info(string.Format("Subscriber 1 received EventMessage with Id {0}.", message.Id));
            Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));

            Logger.Info(string.Format("Message source: {0}.", message.Originator));
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof (ServiceUserUpdateMessageHandler));
    }
}
