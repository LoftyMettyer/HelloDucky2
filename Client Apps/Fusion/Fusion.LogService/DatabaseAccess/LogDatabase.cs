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

        public void AddLogEntry(Guid id, Guid? messageId, string connectorName, Guid? entityRef, DateTime time, char logLevel, string message)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("dbo.pFusionLogAdd",
                    new
                    {
                        Id = id,
                        MessageId = messageId,
                        ConnectorName = connectorName,
                        EntityRef = entityRef,
                        Time = time,
                        LogLevel = logLevel,
                        Message = message
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }

        public void AddTranslationRecord(DateTime time, Guid? messageId, string connectorName, string translationName, Guid busRef, string localId)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("dbo.pIdTranslationLogAdd",
                    new
                    {
                        ConnectorName = connectorName,
                        TranslationName = translationName,
                        BusRef = busRef,
                        LocalId = localId,
                        Time = time,
                        MessageId = messageId
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }
    }
}
