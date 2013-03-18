namespace Connector1.Registries
{
    using System;
    using System.Reflection;
    using StructureMap.Configuration.DSL;
    using ProgressConnector.BusTypeBuilder;
    using Fusion.Messages.General;

    public class BusTypeBuilderRegistry : Registry
    {
        public BusTypeBuilderRegistry()
        {
            For<IBusTypeBuilder>().Use<BusTypeBuilder>();

            Scan(s =>
            {
                s.AssembliesFromApplicationBaseDirectory();
                s.AddAllTypesOf<FusionMessage>()
                    .NameBy(x => x.Name + "FusionMessage");
            });
        }
    }
}
