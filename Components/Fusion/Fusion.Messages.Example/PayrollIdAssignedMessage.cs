using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;
using NServiceBus;

namespace Fusion.Messages.Example
{
    // sole-publisher

    public class PayrollIdAssignedMessage : FusionMessage, IEvent
    {

    }
}
