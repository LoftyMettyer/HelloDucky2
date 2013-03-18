using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NServiceBus;
using StructureMap.Attributes;
using Fusion.Messages.General;
using System.Transactions;

namespace Fusion.Core.Logging
{
    public class FusionLogger : IFusionLogService
    {
        public FusionLogger(string source)
        {
            this.source = source;
        }

        private string source;


        [SetterProperty]
        public IBus Bus
        {
            get;
            set;
        }

        public void LogRefTranslationTransactional(string community, string entityName, Guid busRef, string localId)
        {
            Bus.Send(
                new LogTranslationMessage
                {
                    Id = Guid.NewGuid(),
                    Source = this.source,
                    TimeUtc = DateTime.UtcNow,
                    BusRef = busRef,
                    LocalId = localId,
                    EntityName = entityName,
                    Community = community
                });
        }

     

        public void InfoMessageNonTransactional(string community, Guid messageId, FusionLogLevel logLevel, Guid? entityRef, Guid? primaryEntityRef, string messageName, string message)
        {
            using (TransactionScope ts = new TransactionScope(TransactionScopeOption.RequiresNew))
            {
                Bus.Send(
                    new LogMessage
                    {
                        Id = Guid.NewGuid(),
                        TimeUtc = DateTime.UtcNow,
                        MessageId = messageId,
                        Source = this.source,
                        LogLevel = logLevel,
                        Message = message,
                        MessageName = messageName,
                        EntityRef = entityRef,
                        Community = community,
                        PrimaryEntityRef = primaryEntityRef
                    });

                ts.Complete();
            }
        }

        public void InfoMessageTransactional(string community, Guid messageId, FusionLogLevel logLevel, Guid? entityRef, Guid? primaryEntityRef, string messageName, string message)
        {
            Bus.Send(
                new LogMessage
                {
                    Id = Guid.NewGuid(),
                    TimeUtc = DateTime.UtcNow,
                    MessageId = messageId,
                    Source = this.source,
                    LogLevel = logLevel,
                    Message = message,
                    MessageName = messageName,
                    Community = community,
                    PrimaryEntityRef = primaryEntityRef,
                    EntityRef = entityRef
                });
        }
     
    }
}
