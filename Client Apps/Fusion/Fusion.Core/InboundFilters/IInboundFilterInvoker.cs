namespace Fusion.Core.InboundFilters
{
    using System;
    using Fusion.Messages.General;

    public interface IInboundFilterInvoker
    {
        bool Execute(FusionMessage message);
    }
}
