
namespace Fusion.Core.Test
{
    using System;
    using System.Linq;
    using Fusion.Messages.General;
    using StructureMap;

    public class FusionXmlMetadataExtractInvoker : IFusionXmlMetadataExtractInvoker
    {
        public FusionXmlMetadata GetMetadataFromXml(FusionMessage message)
        {
            var messageType = message.GetType();
            var handlerType = typeof(IFusionXmlMetadataExtract<>).MakeGenericType(messageType);

            var genericHandlerMatches = ObjectFactory.GetAllInstances(handlerType).Cast<IFusionXmlMetadataExtract>();

            if (genericHandlerMatches.Count() != 1)
            {
                throw new ArgumentException("Cannot locate unique IFusionXmlMetadataExtract<T> for type", "message");
            }

            IFusionXmlMetadataExtract sender = genericHandlerMatches.First();

            return sender.GetMetadataFromXml(message);
        }
    }
}
