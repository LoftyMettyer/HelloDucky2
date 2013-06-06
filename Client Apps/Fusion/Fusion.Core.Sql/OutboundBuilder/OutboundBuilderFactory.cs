using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using StructureMap;

namespace Fusion.Core.Sql.OutboundBuilder
{
    public class OutboundBuilderFactor : IOutboundBuilderFactory
    {
        public IOutboundBuilder GetOutboundBuilder(string messageType)
        {
            return ObjectFactory.GetNamedInstance<IOutboundBuilder>(messageType);
        }
    }
}
