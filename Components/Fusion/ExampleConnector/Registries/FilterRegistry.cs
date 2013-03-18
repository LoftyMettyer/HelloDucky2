namespace ExampleConnector.Registries
{
    using System;    
    using StructureMap.Configuration.DSL;
    using Fusion.Core.Conventions;
    using Fusion.Core.InboundFilters;
    using Fusion.Core.OutboundFilters;
   
    public class FilterRegistry : Registry
    {
        public FilterRegistry()
        {
            For<IInboundFilterInvoker>().Use<InboundFilterInvoker>();
            For<IOutboundFilterInvoker>().Use<OutboundFilterInvoker>();

            Scan(s =>
            {
                s.TheCallingAssembly();

                s.Convention<InboundMessageFilters>();
                s.Convention<OutboundMessageFilters>();
            });
        }
    }
}