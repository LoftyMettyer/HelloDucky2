using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.Example;
using Fusion.Core.Sql.OutboundBuilder;
using Fusion.Messages.General;
using Fusion.Core.Sql;
using Connector1.DatabaseAccess;
using Connector1.Configuration;
using Connector1.Messages;

namespace MyPublisher.OutboundBuilders
{
    public class ServiceUserUpdateMessageBuilder : IOutboundBuilder
    {
        public ServiceUserUpdateMessageBuilder(IServiceUserDb serviceUserDb, IBusRefTranslator busRefTranslator, IFusionConfiguration config)
        {
            this.serviceUserDb = serviceUserDb;
            this.refTranslator = busRefTranslator;
            this.config = config;
        }

        private IServiceUserDb serviceUserDb;
        private IBusRefTranslator refTranslator;
        private IFusionConfiguration config;

        public FusionMessage Build(SendFusionMessageRequest source)
        {
            var su = this.serviceUserDb.ReadServiceUser(Convert.ToInt32(source.LocalId));

            if (su == null)
            {
                return null;
            }

            Guid busRef = this.refTranslator.GetBusRef(EntityTranslationNames.ServiceUser, source.LocalId);

            string xml = String.Format("<serviceUserUpdate><ref>{0}</ref><forename>{1}</forename><surname>{2}</surname></serviceUserUpdate>",
                busRef.ToString(), su.Forename, su.Surname
                );

            return new ServiceUserUpdateRequest()
            {
                CreatedUtc = source.TriggerDate,
                Id = Guid.NewGuid(),
                Originator = config.ServiceName,
                EntityRef = busRef,
                Xml = xml
            };

        }
    }
}
