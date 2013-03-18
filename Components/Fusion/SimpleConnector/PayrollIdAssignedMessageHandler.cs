using log4net;

using NServiceBus;
using Fusion.Messages.Example;

namespace Subscriber1
{
    public class PayrollIdAssignedMessageHandler : IHandleMessages<PayrollIdAssignedMessage>
    {
        public void Handle(PayrollIdAssignedMessage message)
        {
            Logger.Info(string.Format("Subscriber 1 received PayrollIdAssignedMessage with Id {0}.", message.Id));
            Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));

            Logger.Info(string.Format("Message source: {0}.", message.Originator));
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(PayrollIdAssignedMessageHandler));
    }
}
