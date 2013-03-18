using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Core.InboundFilters;
using Fusion.Messages.Example;
using log4net;
using Fusion.Core.Sql;
using System.Xml;

namespace Connector1.InboundFilters
{
    public class ServiceUserUpdateLatestInboundMessageFilter : IgnoreNonLatestMessageInboundMessageFilter<ServiceUserUpdateMessage>
    {
     
    }
}
