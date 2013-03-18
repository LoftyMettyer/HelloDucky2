
namespace Fusion.Core.Sql
{
    using System;
    
    public interface IMessageLog
    {
        void AddToMessageLog(string messageType, Guid messageRef, string originator);
        bool IsMessageInLog(string messageType, Guid messageRef);
    }
}
