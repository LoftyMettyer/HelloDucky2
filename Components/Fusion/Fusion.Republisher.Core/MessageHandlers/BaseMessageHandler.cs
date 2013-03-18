using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;
using log4net;
using StructureMap.Attributes;
using Fusion.Core.Sql;
using Fusion.Core.InboundFilters;
using Fusion.Core;
using Fusion.Core.Logging;

namespace Fusion.Publisher.Core.MessageHandlers
{
    public class BaseMessageHandler
    {
        //[SetterProperty]
        //public IFusionConfiguration Configuration
        //{
        //    get;
        //    set;
        //}

        [SetterProperty]
        public IMessageLog MessageLog
        {
            get;
            set;
        }

        [SetterProperty]
        public IMessageTracking MessageTracking
        {
            get;
            set;
        }

        [SetterProperty]
        public IInboundFilterInvoker InboundFilterInvoker
        {
            get;
            set;
        }

        [SetterProperty]
        public IFusionLogService FusionLogger
        {
            get;
            set;
        }



        static protected ILog Logger
        {
            get;
            set;
        }

        public bool StartHandlingMessage(FusionMessage message)
        {
            Logger.Info(string.Format("Connector received " + message.GetMessageName() + " with Id {0} from {1} - xml {2}.", message.Id, message.Originator, message.Xml));

            Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));

            bool shouldProcess = true;

            string messageStatus = null;

            //if (message.Originator == Configuration.ServiceName)
            //{
            //    Logger.Info("We are the originator, ignoring message");
            //    messageStatus = "Ignoring - We are originator";

            //    // We were the originator, we can just ignore
            //    shouldProcess = false;
            //}
            //else
            //{
                // Run inbound filters

                shouldProcess = InboundFilterInvoker.Execute(message);
                if (shouldProcess == false)
                {
                    Logger.Info("Inbound filters indicate message should not be processed");
                    messageStatus = "Inbound filters indicate message should not be processed";
                }
//            }

            FusionLogger.LogMessageReceived(message);
            if (messageStatus != null)
            {
                FusionLogger.InfoMessageTransactional(message, FusionLogLevel.Info, messageStatus);
            }

            MessageLog.AddToMessageLog(message.GetMessageName(), message.Id, message.Originator);

            if (message.EntityRef.HasValue) {
                MessageTracking.SetLastProcessedDate(message.GetMessageName(), message.EntityRef.Value, message.CreatedUtc);
            }

            return shouldProcess;
        }

    }
}
