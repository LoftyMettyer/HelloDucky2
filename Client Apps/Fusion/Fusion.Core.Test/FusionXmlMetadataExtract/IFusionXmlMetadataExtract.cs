
namespace Fusion.Core.Test
{
    using Fusion.Messages.General;

    public interface IFusionXmlMetadataExtract
    {
        FusionXmlMetadata GetMetadataFromXml(FusionMessage message);
    }

    public interface IFusionXmlMetadataExtract<T> where T : FusionMessage
    {
        FusionXmlMetadata GetMetadataFromXml(T message);
    }
}
