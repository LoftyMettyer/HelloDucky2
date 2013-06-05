namespace Fusion.Core.InboundFilters
{
    using System;
    using Fusion.Messages.General;

    public interface IInboundFilterHandler
    {
        bool Handle(FusionMessage message);
    }

    public interface IInboundFilterHandler<TFusionMessage> : IInboundFilterHandler where TFusionMessage : FusionMessage
    {
    }
}