using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

using Dapper;
using System.Data;

namespace Fusion.Core.Sql
{
    public class MessageTracking : IMessageTracking
    {
        public MessageTracking(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        private string ConnectionString
        {
            get;
            set;
        }
        public void SetLastGeneratedDate(string messageType, Guid busRef, DateTime generatedDate)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("fusion.pMessageTrackingSetLastGeneratedDate",
                    new
                    {
                        MessageType = messageType,
                        LastGeneratedDate = generatedDate.ToUniversalTime(),
                        BusRef = busRef
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }

        public void SetLastProcessedDate(string messageType, Guid busRef, DateTime processedDate)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("fusion.pMessageTrackingSetLastProcessedDate",
                    new
                    {
                        MessageType = messageType,
                        LastProcessedDate = processedDate.ToUniversalTime(),
                        BusRef = busRef
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }

        public MessageTimes GetMessageTimes(string messageType, Guid busRef)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var r = c.Query<MessageTimes>("fusion.pMessageTrackingGetLastMessageDates",
                            new
                            {
                                MessageType = messageType,
                                BusRef = busRef
                            },
                            commandType: CommandType.StoredProcedure);

                MessageTimes messageTimes = r.FirstOrDefault();

                if (messageTimes == null)
                {
                    return new MessageTimes();
                }

                return messageTimes;
            }
        }

        public void SetLastGeneratedXml(string messageType, Guid busRef, string xml)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("fusion.pMessageTrackingSetLastGeneratedXml",
                    new
                    {
                        MessageType = messageType,
                        LastGeneratedXml = new DbString { Value = xml, Length = -1 },
                        BusRef = busRef
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }

        public string GetLastGeneratedXml(string messageType, Guid busRef)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var r = c.Query<string>("fusion.pMessageTrackingGetLastGeneratedXml",
                            new
                            {
                                MessageType = messageType,
                                BusRef = busRef
                            },
                            commandType: CommandType.StoredProcedure);

                string lastGeneratedXml = r.FirstOrDefault();

                return lastGeneratedXml;
            }
        }

    }
}
