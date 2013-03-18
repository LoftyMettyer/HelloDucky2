using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.LogService.DatabaseAccess
{
    public class NullLogDatabase : ILogDatabase
    {

        public void AddLogEntry(Guid id, string messageName, Guid? messageId, string community, string connectorName, Guid? entityRef, Guid? primaryEntityRef, DateTime time, char logLevel, string message)
        {
        }

        public void AddTranslationRecord(DateTime time, Guid? messageId, string connectorName, string community, string translationName, Guid busRef, string localId)
        {
        }


        public void AddMessageAudit(string queueMessageId, string queueOriginalMessageId, string queueHeaders, string queueMessageType, DateTime queueReceiveTime, string queueReplyToAddress, Guid messageId, string messageOriginator, Guid? entityRef, Guid? primaryEntityRef, string community, string xml, DateTime createdUtc, int schemaVersion)
        {
        }
    }
}
