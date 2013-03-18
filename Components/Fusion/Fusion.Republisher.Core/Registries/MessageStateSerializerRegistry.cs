using Fusion.Core.Logging;
using Fusion.Republisher.Core.MessageStateSerializer;
using StructureMap.Configuration.DSL;

namespace Fusion.Republisher.Core.Registries
{
    public class MessageStateSerializerRegistry : Registry
    {
        public MessageStateSerializerRegistry()
        {
            For<IMessageStateSerializer>().Use<JsonMessageStateSerializer>();
        }
    }
}
