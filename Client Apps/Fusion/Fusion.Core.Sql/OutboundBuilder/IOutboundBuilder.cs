using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;

namespace Fusion.Core.Sql.OutboundBuilder
{

    public interface IOutboundBuilder
    {
        FusionMessage Build(SendFusionMessageRequest source);
    }
}
