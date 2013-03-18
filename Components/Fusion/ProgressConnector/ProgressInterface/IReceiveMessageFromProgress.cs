using System;
namespace ProgressConnector.ProgressInterface
{
    public interface IReceiveMessageFromProgress
    {
        RawOpenExchangeData ReceiveOneMessage();
        void AcknowledgeSent(Guid id);
    }
}
