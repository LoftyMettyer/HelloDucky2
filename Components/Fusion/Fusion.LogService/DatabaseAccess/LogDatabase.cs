using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using Dapper;
using System.Data;

namespace Fusion.LogService.DatabaseAccess
{
    public class LogDatabase : ILogDatabase
    {
        /// <summary>
        /// Initializes a new instance of the BusRefTranslator class.
        /// </summary>
        /// <param name="connectionString"> The connection string for the database. </param>
        public LogDatabase(string connectionString)
        {
            this.ConnectionString = connectionString;
        }


        /// <summary>
        /// Gets or sets the connection string.
        /// </summary>
        /// <value>
        /// The connection string.
        /// </value>
        private string ConnectionString
        {
            get;
            set;
        }

        public void AddLogEntry(Guid id, string messageName, Guid? messageId, string community, string connectorName, Guid? entityRef, Guid? primaryEntityRef, DateTime time, char logLevel, string message)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("dbo.pFusionLogAdd",
                    new
                    {
                        Id = id,
                        MessageName = messageName,
                        MessageId = messageId,
                        Community = community,
                        ConnectorName = connectorName,
                        EntityRef = entityRef,
                        PrimaryEntityRef = primaryEntityRef,
                        Time = time,
                        LogLevel = logLevel,
                        Message = message
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }

        public void AddTranslationRecord(DateTime time, Guid? messageId, string connectorName, string community, string translationName, Guid busRef, string localId)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("dbo.pIdTranslationLogAdd",
                    new
                    {
                        ConnectorName = connectorName,
                        Community = community,
                        TranslationName = translationName,
                        BusRef = busRef,
                        LocalId = localId,
                        Time = time,
                        MessageId = messageId
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }


        public void AddMessageAudit(string queueMessageId, string queueOriginalMessageId, string queueHeaders, string queueMessageType, DateTime queueReceiveTime, string queueReplyToAddress, Guid messageId, string messageOriginator, Guid? entityRef, Guid? primaryEntityRef, string community, string xml, DateTime createdUtc, int schemaVersion)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("dbo.pMessageAuditAdd",
                    new
                    {
                        QueueMessageId = queueMessageId,
                        QueueOriginalMessageId = queueOriginalMessageId,
                        QueueHeaders = queueHeaders,
                        QueueMessageType = queueMessageType,
                        QueueReceiveTime = queueReceiveTime,
                        QueueReplyToAddress = queueReplyToAddress,
                        MessageId = messageId,
                        MessageOriginator = messageOriginator,
                        EntityRef = entityRef,
                        PrimaryEntityRef = primaryEntityRef,
                        Community = community,
                        Xml = xml,
                        CreatedUtc = createdUtc,
                        SchemaVersion = schemaVersion                       
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }


    }
}
