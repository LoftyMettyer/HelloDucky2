namespace ExampleConnector.Registries
{
    using System;
    using System.Reflection;
    using StructureMap.Configuration.DSL;
    using Fusion.Core.Sql.OutboundBuilder;

    public class OutboundBuilderRegistry : Registry
    {
        public OutboundBuilderRegistry()
        {
            For<IOutboundBuilderFactory>().Use<OutboundBuilderFactor>();

            Scan(s =>
            {
                s.Assembly(Assembly.GetExecutingAssembly());
                s.AddAllTypesOf<IOutboundBuilder>()
                    .NameBy(x => x.Name.Replace("MessageBuilder", String.Empty));
            });
        }
    }
}
