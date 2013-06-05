namespace Fusion.Core.OutboundFilters
{
    using System;
    using Fusion.Messages.General;

    public interface IOutboundFilterHandler
    {
        bool Handle(FusionMessage message);
    }

    public interface IOutboundFilterHandler<TFusionMessage> : IOutboundFilterHandler where TFusionMessage : FusionMessage
    {
        bool Handle(TFusionMessage message);
    }
}