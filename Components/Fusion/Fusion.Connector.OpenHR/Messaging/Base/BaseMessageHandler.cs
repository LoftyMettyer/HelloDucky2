using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Connector.OpenHR.Messaging;
using Fusion.Messages.General;
using NServiceBus.Encryption;
using log4net;
using StructureMap.Attributes;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Core.Sql;
using Fusion.Core.InboundFilters;
using Fusion.Core;
using Fusion.Core.Logging;

namespace Fusion.Connector.OpenHR.MessageHandlers
{
    public class BaseMessageHandler
    {
        [SetterProperty]
        public IFusionConfiguration Configuration
        {
            get;
            set;
        }

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

        [SetterProperty]
        static protected ILog Logger
        {
            get;
            set;
        }

        [SetterProperty]
        public IBusRefTranslator BusRefTranslator { get; set; }

        protected readonly string ConnectionString;

        public BaseMessageHandler()
        {
            ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            Logger = LogManager.GetLogger(typeof(BaseMessageHandler));
        }

        public bool StartHandlingMessage(FusionMessage message)
        {
            Logger.Info(string.Format("Connector received " + message.GetMessageName() + " with Id {0} from {1} - xml {2}.", message.Id, message.Originator, message.Xml));

            Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));

            bool shouldProcess = true;

            string messageStatus = null;

            if (message.Originator == Configuration.ServiceName)
            {
                Logger.Info("We are the originator, ignoring message");
                messageStatus = "Ignoring - We are originator";

                // We were the originator, we can just ignore
                shouldProcess = false;
            }
            else
            {
                // Run inbound filters

                shouldProcess = InboundFilterInvoker.Execute(message);
                if (shouldProcess == false)
                {
                    Logger.Info("Inbound filters indicate message should not be processed");
                    messageStatus = "Inbound filters indicate message should not be processed";
                }
            }

            FusionLogger.LogMessageReceived(message);
            if (messageStatus != null)
            {
                FusionLogger.InfoMessageTransactional(message, FusionLogLevel.Info, messageStatus);
            }

            MessageLog.AddToMessageLog(message.GetMessageName(), message.Id, message.Originator);

            if (message.EntityRef.HasValue)
            {
                MessageTracking.SetLastProcessedDate(message.GetMessageName(), message.EntityRef.Value, message.CreatedUtc);
            }

						if (IsAlreadyprocessed(message))
						{
							shouldProcess = false;
						}

            return shouldProcess;
        }

				public bool XMLCompare(object a, object b)
				{
					// They're both null.
					if (a == null && b == null) return true;
					// One is null, so they can't be the same.
					if (a == null || b == null) return false;
					// How can they be the same if they're different types?
					if (a.GetType() != b.GetType()) return false;
					var Props = a.GetType().GetProperties();
					foreach (var Prop in Props)
					{
						// See notes *
						var aPropValue = Prop.GetValue(a) ?? string.Empty;
						var bPropValue = Prop.GetValue(b) ?? string.Empty;
						if (aPropValue.ToString() != bPropValue.ToString())
							return false;
					}
					return true;
				}

				private bool IsAlreadyprocessed(FusionMessage message)
				{
					var messageType = message.GetMessageName();

					if (message.EntityRef != null)
					{
						var busRef = message.EntityRef.Value;

						var lastMessageGenerated = MessageTracking.GetLastGeneratedXml(messageType, busRef);

						// Simple compare, could be more elaborate
						if (lastMessageGenerated == message.Xml)
						{
							Logger.InfoFormat("Inbound Message {0} {1} no changes", messageType, busRef);
							return true;
						}
					}

					return false;
				}

    }
}
