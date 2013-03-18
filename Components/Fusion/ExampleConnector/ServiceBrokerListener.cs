using System;
using System.Transactions;
using Fusion.Core.MessageSenders;
using Fusion.Core.Sql.OutboundBuilder;
using Fusion.Core.Sql.ServiceBroker;
using Fusion.Messages.General;
using NServiceBus;
using StructureMap.Attributes;
using Fusion.Core.OutboundFilters;
using log4net;
using Fusion.Core.Logging;

namespace MyPublisher
{
    public class ServiceBrokerListenerStartup : IWantToRunAtStartup
    {
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
        public IOutboundBuilderFactory OutboundBuilderFactory
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
        public IOutboundFilterInvoker OutboundFilterInvoker
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


        private static readonly ILog Logger = LogManager.GetLogger(typeof(ServiceBrokerListenerStartup));

        public void Run()
        {
            Logger.Info("Listening to service broker..." );

            var sl = this.ServiceBrokerListener;

            for (; ; )
            {
                using (TransactionScope ts = new TransactionScope(TransactionScopeOption.Required, new TransactionOptions
                {
                    IsolationLevel = IsolationLevel.ReadCommitted
                }))
                {
                    var message = sl.ReceiveMessage();

                    if (message != null)
                    {
                        var thisBuilder = OutboundBuilderFactory.GetOutboundBuilder(message.MessageType);
                        var fusionMessage = thisBuilder.Build(message);

                        // Builder can return null indicating that no message needs to be sent, for whatever it's reason

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
                    

                    ts.Complete();
                }
            }

        }

        public void Stop()
        {

        }
    }
}
