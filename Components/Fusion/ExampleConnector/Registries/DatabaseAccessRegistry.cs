using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using StructureMap.Configuration.DSL;
using Connector1.DatabaseAccess;
using Fusion.Core.Sql;
using Fusion.Core.Sql.ServiceBroker;
using StructureMap;
using Fusion.Core.Logging;

namespace ExampleConnector.Registries
{
    public class DatabaseAccessRegistry : Registry
    {
        public DatabaseAccessRegistry()
        {
            string connectionString = "user id=sa;password=asr;initial catalog=FusionExample;data source=.;Connect Timeout=30;Application Name=Fusion Connector";
            string community = "exampleCommunity";

            For<IServiceUserDb>().Use<ServiceUserDb>().Ctor<string>("connectionString").Is(connectionString);

            /* Core fusion sql resources we are going to use */
            For<IMessageTracking>().Use<MessageTracking>().Ctor<string>("connectionString").Is(connectionString);
            For<IBusRefTranslator>().Use<BusRefTranslator>()
                .EnrichWith<IBusRefTranslator>((c, x) => new LoggingBusRefTranslatorDecorator(community, x, c.GetInstance<IFusionLogService>()))
                .Ctor<string>("connectionString").Is(connectionString);
                
//                ).Ctor<IBusRefTranslator>("busTranslator").Is<BusRefTranslator>().Ctor<string>("connectionString").Is(connectionString);
            For<IFusionServiceBrokerListener>().Use<FusionServiceBrokerListener>().Ctor<string>("connectionString").Is(connectionString);
            For<IMessageLog>().Use<MessageLog>().Ctor<string>("connectionString").Is(connectionString);
        } 
    }
}
