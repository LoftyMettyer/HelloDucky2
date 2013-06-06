
namespace Fusion.Core.Test
{
    using Fusion.Messages.General;
    using System;

    public abstract class FusionXmlMetadataExtract<T> : IFusionXmlMetadataExtract, IFusionXmlMetadataExtract<T> where T : FusionMessage
    {
        public FusionXmlMetadata GetMetadataFromXml(FusionMessage message)
        {
            return GetMetadataFromXml((T)message);
        }

        public abstract FusionXmlMetadata GetMetadataFromXml(T message);        
    }
}
