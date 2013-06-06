using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Core.Sql.OutboundBuilder
{
    public interface IOutboundBuilderFactory
    {
        IOutboundBuilder GetOutboundBuilder(string messageType);
    }
}
