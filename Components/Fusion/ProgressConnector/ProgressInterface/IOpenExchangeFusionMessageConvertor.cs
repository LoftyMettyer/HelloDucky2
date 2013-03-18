using System;
using Fusion.Messages.General;
namespace Connector1.ProgressInterface
{
    public interface IOpenExchangeFusionMessageConvertor
    {
        string BuildOpenExchangeMessage(FusionMessage message);
    }
}
