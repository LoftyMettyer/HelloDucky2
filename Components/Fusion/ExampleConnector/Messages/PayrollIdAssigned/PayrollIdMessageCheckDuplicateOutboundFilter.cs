using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.Example;
using Fusion.Core.OutboundFilters;
using Fusion.Core.MessageValidators;
using log4net;
using Fusion.Core;
using Fusion.Core.Sql;

namespace Connector1.OutboundFilters
{
    public class PayrollIdAssignedCheckDuplicateOutboundFilter : CheckDuplicateOutboundFilterHandler<PayrollIdAssignedMessage>
    {

    }
}
