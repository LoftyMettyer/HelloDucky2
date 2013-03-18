
namespace Fusion.Core.Test
{
    using System;
    using Fusion.Core.MessageSenders;

    public interface IOutboundMessageWatcher
    {
        void Start();
        void Stop();
        IMessageSender MessageSender { get; set; }
    }
}
