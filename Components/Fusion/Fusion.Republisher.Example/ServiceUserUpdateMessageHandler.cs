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
    public class ServiceUserUpdateEventHandler : IHandleMessages<ServiceUserUpdateRequest> {

        public IBus Bus
        {
            get;
            set;
        }

        public void Handle(ServiceUserUpdateRequest message)
        {
            Logger.Info("Hello");
            Logger.Info(string.Format("Fusion Publisher received ServiceUserUpdate with Id {0}.", message.Id));
            Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));

            Logger.Info(string.Format("Message name: {0}.", message.GetType()));
            Logger.Info(string.Format("Message source: {0}.", message.Originator));

            Logger.Info("Republishing to bus");

            Bus.Publish(new ServiceUserUpdateMessage
            {
                CreatedUtc = message.CreatedUtc,
                EntityRef = message.EntityRef,
                Id = message.Id,
                Originator = message.Originator,
                SchemaVersion = message.SchemaVersion,
                Xml = message.Xml
            }
           );
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(ServiceUserUpdateEventHandler));
    }
    
    
}
