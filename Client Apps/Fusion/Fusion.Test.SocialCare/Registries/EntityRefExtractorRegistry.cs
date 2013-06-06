

namespace Fusion.Test.Registries
{
    using System.Reflection;
    using Fusion.Core.Test;
    using StructureMap.Configuration.DSL;

    public class FusionXmlMetadataExtractorRegistry : Registry
    {
        public FusionXmlMetadataExtractorRegistry()
        {
            For<IFusionXmlMetadataExtractInvoker>().Use<FusionXmlMetadataExtractInvoker>();

            Scan(
                s =>
                {
                    s.Assembly(Assembly.GetExecutingAssembly());
                    s.ConnectImplementationsToTypesClosing(typeof(IFusionXmlMetadataExtract<>));
                }

                );
        }
    }
}
