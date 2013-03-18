using System;
using System.IO;

namespace Fusion.Core.Sql
{
    public interface ISendFusionMessageRequestBuilder
    {
        SendFusionMessageRequest Build(Stream bodyStream);
    }
}
