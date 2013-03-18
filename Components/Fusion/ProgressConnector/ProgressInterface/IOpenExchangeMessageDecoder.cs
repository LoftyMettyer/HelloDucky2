using System;
namespace ProgressConnector.ProgressInterface
{
    public interface IOpenExchangeMessageDecoder
    {
        OpenExchangeGeneratedContent Decode(RawOpenExchangeData data);
    }
}
