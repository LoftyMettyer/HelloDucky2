namespace Fusion.Core.OutboundFilters
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Fusion.Messages.General;

    public abstract class OutboundFilterHandler<TFusionMessage> : IOutboundFilterHandler<TFusionMessage> where TFusionMessage : FusionMessage
    {
        public bool Handle(FusionMessage message)
        {
            return this.Handle((TFusionMessage)message);
        }

        public abstract bool Handle(TFusionMessage message);
    }
}
