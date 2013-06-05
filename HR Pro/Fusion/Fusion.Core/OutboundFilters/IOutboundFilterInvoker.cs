namespace Fusion.Core.OutboundFilters
{
    using System;
    using Fusion.Messages.General;

    public interface IOutboundFilterInvoker
    {
        bool Execute(FusionMessage message);
    }
}
