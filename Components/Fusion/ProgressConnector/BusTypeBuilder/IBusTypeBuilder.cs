using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;
using Connector1.ProgressInterface;
using ProgressConnector.ProgressInterface;

namespace ProgressConnector.BusTypeBuilder
{

    public interface IBusTypeBuilder
    {
        FusionMessage Build(OpenExchangeMessage source);
    }
}
