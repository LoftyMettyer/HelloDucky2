
namespace Fusion.Core.Test
{
    using Fusion.Messages.General;

    public interface IFusionXmlMetadataExtractInvoker
    {
        FusionXmlMetadata GetMetadataFromXml(FusionMessage message);
    }
}
