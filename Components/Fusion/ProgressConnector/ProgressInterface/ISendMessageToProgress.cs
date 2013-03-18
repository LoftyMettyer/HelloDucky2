
namespace ProgressConnector.ProgressInterface
{
    using System;
    using Fusion.Messages.General;
    
    public interface ISendMessageToProgress
    {
        ProgressSendStatus SendMessage(FusionMessage message);
    }
}
