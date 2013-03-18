using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;
using NServiceBus;

namespace Fusion.Messages.Example
{
    // Multi-master

    public class CostCentreChangeRequest : FusionMessage, ICommand
    {
    }


    public class CostCentreChangeMessage : FusionMessage, IEvent
    {
    }
}
