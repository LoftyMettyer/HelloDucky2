using System;
using System.Data.SqlClient;
using System.Transactions;
using Fusion.Core.MessageSenders;
using Fusion.Core.Sql;
using Fusion.Core.Sql.OutboundBuilder;
using Fusion.Core.Sql.ServiceBroker;
using Fusion.Messages.General;
using NServiceBus;
using StructureMap.Attributes;
using Fusion.Core.OutboundFilters;
using log4net;
using Fusion.Core.Logging;
using Fusion.Core;

namespace Fusion
{
    public class ServiceBrokerListenerStartup : IWantToRunAtStartup
    {
        [SetterProperty]
        public IBus Bus
        {
            get;
            set;
        }

        [SetterProperty]
        public IMessageSenderInvoker MessageSenderInvoker
        {
            get;
            set;
        }

        [SetterProperty]
        public IFusionServiceBrokerListener ServiceBrokerListener
        {
            get;
            set;
        }

        [SetterProperty]
        public IOutboundBuilderFactory OutboundBuilderFactory
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
        public IMessageTracking MessageTracking
        {
            get;
            set;
        }

        [SetterProperty]
        public IOutboundFilterInvoker OutboundFilterInvoker
        {
            get;
            set;
        }


        protected void TrackMessage(FusionMessage message)
        {
            MessageTracking.SetLastGeneratedDate(message.GetMessageName(), message.EntityRef.Value, message.CreatedUtc);
            MessageTracking.SetLastGeneratedXml(message.GetMessageName(), message.EntityRef.Value, message.Xml);
        }


        protected bool LaterInboundMessageProcessed(FusionMessage message)
        {
            MessageTimes times = MessageTracking.GetMessageTimes(message.GetMessageName(), message.EntityRef.Value);

            if (times.LastProcessedDate.HasValue && times.LastProcessedDate > message.CreatedUtc)
            {
                return true;
            }

            return false;
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(ServiceBrokerListenerStartup));

        public void Run()
        {
            Logger.Info("Listening to service broker...");

            var sl = this.ServiceBrokerListener;

            for (; ; )
            {
                using (var ts = new TransactionScope(TransactionScopeOption.Required, new TransactionOptions{ IsolationLevel = IsolationLevel.ReadCommitted })) 
				{
					try {
						var message = sl.ReceiveMessage();

                        if (message != null)
                        {

                            var thisBuilder = OutboundBuilderFactory.GetOutboundBuilder(message.MessageType);
                            var fusionMessage = thisBuilder.Build(message);

                            if (fusionMessage != null)
                            {
                                FusionLogger.LogMessageGenerated(fusionMessage);

                                // Execute outbound filters
                                bool canSend = OutboundFilterInvoker.Execute(fusionMessage);

                                if (canSend)
                                {
                                    MessageSenderInvoker.Invoke(fusionMessage);
                                }
                                else
                                {
                                    Logger.InfoFormat("Outbound message {0}/{1} execution stopped by filters", message.MessageType, message.LocalId);
                                    FusionLogger.InfoMessageTransactional(fusionMessage, FusionLogLevel.Info, "Message stopped by outbound filters");
                                }
                            }


                        }
					} 
					catch (SqlException se) {
						Logger.WarnFormat("Unable to receive message from service broker: {0}", se.Message);
						System.Threading.Thread.Sleep(5000);
					}
					catch (Exception e)
					{
						Logger.ErrorFormat("Unable to process service broker message: {0}", e.Message);
					}

                    ts.Complete();
                }
            }

        }

        public void Stop()
        {

        }
    }
}
