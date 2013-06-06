using System;

namespace Fusion.Core.Sql
{
    public interface IMessageTracking
    {
        string GetLastGeneratedXml(string messageType, Guid busRef);
        MessageTimes GetMessageTimes(string messageType, Guid busRef);
        void SetLastGeneratedDate(string messageType, Guid busRef, DateTime generatedDate);
        void SetLastGeneratedXml(string messageType, Guid busRef, string xml);
        void SetLastProcessedDate(string messageType, Guid busRef, DateTime processedDate);
    }
}
