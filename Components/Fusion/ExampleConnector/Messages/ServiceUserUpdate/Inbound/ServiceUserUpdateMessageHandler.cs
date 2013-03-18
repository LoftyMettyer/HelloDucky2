using log4net;

using NServiceBus;
using Fusion.Messages.Example;
using Connector1.DatabaseAccess;
using StructureMap.Attributes;
using System.Xml;
using Connector1.Configuration;
using System.Xml.Linq;
using System.IO;
using System;
using Fusion.Core.Sql;
using Fusion.Core.MessageValidators;
using Fusion.Core;
using Fusion.Core.InboundFilters;
using Connector1.Messages;

namespace Connector1.MessageHandlers
{
    public class ServiceUserUpdateMessageHandler : BaseMessageHandler, IHandleMessages<ServiceUserUpdateMessage>
    {
        [SetterProperty]
        public IServiceUserDb ServiceUserDb
        {
            get;
            set;
        }

        [SetterProperty]
        public IBusRefTranslator BusRefTranslator
        {
            get;
            set;
        }

        public void Handle(ServiceUserUpdateMessage message)
        {
            bool shouldProcess = base.StartHandlingMessage(message);


            if (shouldProcess == false) return;

            // processing perhaps should not be in here? but rather in a deconstruction class

            XDocument loaded = XDocument.Load(new StringReader(message.Xml));

            //<serviceUserUpdate><ref>{0}</ref><forename>{1}</forename><surname>{2}</surname></serviceUserUpdate>

            var rootNode = loaded.Element("serviceUserUpdate");

            string r = (string)rootNode.Element("ref");
            string forename = (string)rootNode.Element("forename");
            string surname = (string)rootNode.Element("surname");

            Guid busRef = new Guid(r);

            string localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.ServiceUser, busRef);

            // Set fusion context on database updates to prevent ServiceBroker re-firing messages

            ServiceUserDb.MessageContext = "ServiceUserUpdate";

            if (localId == null)
            {
                int newId = ServiceUserDb.CreateServiceUser(forename, surname);
                BusRefTranslator.SetBusRef(EntityTranslationNames.ServiceUser, newId.ToString(), busRef);
            }
            else
            {
                ServiceUserDb.UpdateServiceUser(Convert.ToInt32(localId), forename, surname);
            }

        }

        static ServiceUserUpdateMessageHandler()
        {
            Logger = LogManager.GetLogger(typeof(ServiceUserUpdateMessageHandler));
        }

    }
}
