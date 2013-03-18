using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.Example;
using Fusion.Core.OutboundFilters;
using Fusion.Core.MessageValidators;
using log4net;
using Fusion.Core;

namespace Connector1.OutboundFilters
{
    public class ServiceUserUpdateMessageSchemaValidatorOutboundFilter : SchemaValidatorOutboundFilterHandler<ServiceUserUpdateRequest>
    {
        public override bool Handle(ServiceUserUpdateRequest message)
        {
            bool valid = base.CheckValidity(message);
           
            // Would now return false to not send outbound message
            return true;
        }
    }
}
