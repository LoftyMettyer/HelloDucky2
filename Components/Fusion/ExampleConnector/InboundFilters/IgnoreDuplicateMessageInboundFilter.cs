

namespace Connector1.InboundFilters
{
    using Fusion.Core;
    using Fusion.Core.InboundFilters;
    using Fusion.Core.Sql;
    using Fusion.Messages.General;
    using log4net;
    using StructureMap.Attributes;

    public abstract class IgnoreDuplicateMessageInboundFilter<T> : InboundFilterHandler<T> where T : FusionMessage
    {
        [SetterProperty]
        public IMessageLog MessageLog
        {
            get;
            set;
        }

        protected ILog Logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public override bool Handle(T message)
        {
            string messageType = message.GetMessageName();

            bool isInMessageLog = MessageLog.IsMessageInLog(messageType, message.Id);

            Logger.InfoFormat("IgnoreDuplicateMessage - {0}/{1} - returns {2}", messageType, message.Id, isInMessageLog);

            return !isInMessageLog;
        }

    }
}
