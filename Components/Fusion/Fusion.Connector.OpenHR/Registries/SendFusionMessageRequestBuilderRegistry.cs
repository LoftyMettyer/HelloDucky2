using Fusion.Core.Sql;
using StructureMap.Configuration.DSL;

namespace Fusion.Connector.OpenHR.Registries
{
    public class SendFusionMessageRequestBuilderRegistry : Registry
    {
        public SendFusionMessageRequestBuilderRegistry()
        {
            For<ISendFusionMessageRequestBuilder>().Use<SendFusionMessageRequestBuilder>();
        }
    }
}
