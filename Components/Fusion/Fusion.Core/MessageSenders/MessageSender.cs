
namespace Fusion.Core.MessageSenders
{
    using Fusion.Messages.General;

    public abstract class MessageSender<T> : IMessageSender, IMessageSender<T> where T : FusionMessage
    {
        public void Send(FusionMessage message)
        {
            Send((T)message);
        }

        public abstract void Send(T message);        
    }
}
