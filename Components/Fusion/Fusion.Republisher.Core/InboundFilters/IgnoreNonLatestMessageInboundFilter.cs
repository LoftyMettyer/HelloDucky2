using System;
using Fusion.Core;
using Fusion.Core.InboundFilters;
using Fusion.Core.Sql;
using Fusion.Messages.General;
using log4net;
using StructureMap.Attributes;

namespace Fusion.Publisher.Core.InboundFilters
{
    public abstract class IgnoreNonLatestMessageInboundMessageFilter<T> : InboundFilterHandler<T> where T: FusionMessage
    {
        [SetterProperty]
        public IMessageTracking MessageTracking
        {
            get;
            set;
        }

        protected ILog Logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public override bool Handle(T message)
        {
            string messageType = message.GetMessageName();

            Guid busRef = message.EntityRef.Value;

            MessageTimes messageTimes = this.MessageTracking.GetMessageTimes(messageType, busRef);

            bool shouldProcess = true;

            if (messageTimes.LastGeneratedDate.HasValue)
            {
                if (message.CreatedUtc < messageTimes.LastGeneratedDate.Value)
                {
                    shouldProcess = false;
                }
            }

            if (messageTimes.LastProcessedDate.HasValue)
            {
                if (message.CreatedUtc < messageTimes.LastProcessedDate.Value)
                {
                    shouldProcess = false;
                }
            }

            if (shouldProcess == false)
            {
                Logger.InfoFormat("Message {0}/{1} for {2} dated {3} is older than last message seen locally ", messageType, message.Id, busRef, message.CreatedUtc);

                return false;
            }

            return true;
        }
    }
}
