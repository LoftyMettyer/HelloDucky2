
namespace Fusion.Core.Sql.ServiceBroker
{
    using System.Data.SqlClient;
    using global::ServiceBroker.Net;
    using log4net;

    public class FusionServiceBrokerListener : IFusionServiceBrokerListener
    {
        public FusionServiceBrokerListener(string connectionString, ISendFusionMessageRequestBuilder requestBuilder)
        {
            this.ConnectionString = connectionString;
            this.SendFusionMessageRequestBuilder = requestBuilder;
        }

        const int MessageTimeout = 10 * 1000;

        const string InitiatorServiceName = "FusionApplicationService";
        const string TargetServiceName = "FusionConnectorService";
        const string MessageContractName = "TriggerFusionContract";
        const string MessageType = "TriggerFusionSend";

        const string QueueName = "fusion.qFusion";

        private static readonly ILog Logger = LogManager.GetLogger(typeof(FusionServiceBrokerListener));

        public string ConnectionString
        {
            get;
            private set;
        }

        private ISendFusionMessageRequestBuilder SendFusionMessageRequestBuilder
        {
            get;
            set;
        }


        public SendFusionMessageRequest ReceiveMessage()
        {
            // We are using a Transaction Scope so that all relevant activities that can enlist
            // in a distributed transaction can be covered.
            // 
            // This includes bus transmission, retrieval of messages from service broker,
            // and any updates

            SendFusionMessageRequest result = null;

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                Logger.Debug("Waiting for service broker message");

                try
                {
                    Message m = ServiceBrokerWrapperImplicitTransaction.WaitAndReceive(conn, QueueName, MessageTimeout);

                    if (m != null)
                    {

                        if (m.MessageTypeName != Message.EndDialogType)
                        {
                            // Do something

                            Logger.Info("Received Message from service broker " + m.ServiceName + " " + m.ServiceContractName + " " + m.MessageTypeName);

                            if (m.Body.Length > 0)
                            {
                                var sfmr = this.SendFusionMessageRequestBuilder.Build(m.BodyStream);

                                result = sfmr;
                                Logger.Info("Service Broker message requests send " + sfmr.MessageType + " id " + sfmr.LocalId);
                            }

                            ServiceBrokerWrapperImplicitTransaction.EndConversation(conn, m.ConversationHandle);
                        }
                    }
                }
                catch (SqlException s)
                {
                    Logger.Error("Sql Exception from Service Broker", s);
                    if (s.Number == 9617)
                    {
                        Logger.Error("Service Queue is disabled (probably a poison message handling) - re-enable with ALTER QUEUE " + QueueName + " WITH STATUS = ON;");
                    }

                    throw;
                }
            }

            return result;
        }
    }
}
