using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MyPublisher;
using NServiceBus;
using log4net;
using Fusion.Messages.Example;

namespace Subscriber1
{
    public class CostCentreChangeMessageHandler : IHandleMessages<CostCentreChangeMessage>
    {

        public IBus Bus
        {
            get;
            set;
        }

        public void Handle(CostCentreChangeMessage message)
        {
            Logger.Info(string.Format("Fusion Publisher received CostCentreChangeMessage with Id {0}.", message.Id));
            Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));

            Logger.Info(string.Format("Message name: {0}.", message.GetType()));
            Logger.Info(string.Format("Message source: {0}.", message.Originator));

            Logger.Info("Republishing to bus");

            Bus.Publish(message);
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(CostCentreChangeMessageHandler));
    }


}
