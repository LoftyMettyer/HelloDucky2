using System;
using System.Data;
using System.Data.SqlClient;
using Dapper;

namespace Fusion.Core.Sql
{
    public class MessageLog : IMessageLog 
    {
        public MessageLog(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        private string ConnectionString
        {
            get;
            set;
        }

        public void AddToMessageLog(string messageType, Guid messageRef, string originator)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("fusion.pMessageLogAdd",
                    new
                    {
                        MessageType = messageType,
                        MessageRef = messageRef,
                        Originator = originator
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }

        public bool IsMessageInLog(string messageType, Guid messageRef)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var p = new DynamicParameters();
                p.Add("@MessageType", messageType);
                p.Add("@MessageRef", messageRef);
                p.Add("@ReceivedBefore", dbType: DbType.Boolean, direction: ParameterDirection.Output);

                c.Execute("fusion.pMessageLogCheck",
                   p,
                    commandType: CommandType.StoredProcedure);

                return p.Get<bool>("@ReceivedBefore");
            }
        }

    }
}
