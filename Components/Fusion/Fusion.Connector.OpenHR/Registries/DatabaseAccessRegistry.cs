using Fusion.Connector.OpenHR.Messaging;
using Fusion.Core.Sql.OutboundBuilder;
using StructureMap.Configuration.DSL;
using Fusion.Core.Sql;
using Fusion.Core.Sql.ServiceBroker;
using Fusion.Core.Logging;
using System.Configuration;

namespace Fusion.Connector.OpenHR.Registries
{
    public class DatabaseAccessRegistry : Registry
    {
        public DatabaseAccessRegistry()
        {
            string dbName = ConfigurationManager.AppSettings["OpenHR_db"];
            string serverName = ConfigurationManager.AppSettings["OpenHR_server"];

            string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            string community = ConfigurationManager.AppSettings["Community"];

            //connectionString = string.Format("Data Source={0};Initial Catalog={1};Integrated Security=True;APP=OpenHR Fusion Connector", serverName, dbName);


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
