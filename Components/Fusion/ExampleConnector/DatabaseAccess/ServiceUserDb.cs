using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Dapper;
using System.Data;
using System.Data.SqlClient;

namespace Connector1.DatabaseAccess
{
    public class ServiceUserDb : IServiceUserDb
    {

        public ServiceUserDb(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        private string ConnectionString
        {
            get;
            set;
        }

        public string MessageContext
        {
            get;
            set;
        }

        public void UpdateServiceUser(int id, string forename, string surname)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                if (MessageContext != null)
                {
                    c.Execute("fusion.pSetFusionContext", new
                    {
                        MessageType = MessageContext
                    },
                    commandType: CommandType.StoredProcedure);
                }

                c.Execute(@"
  update ServiceUsers
     set Forename = @Forename,
         Surname = @Surname
     where ServiceUserId = @ServiceUserId
",
                    new
                    {
                        Forename = forename,
                        Surname = surname,
                        ServiceUserId = id
                    });
            }
        }

        public int CreateServiceUser(string forename, string surname)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                if (MessageContext != null)
                {
                    c.Execute("fusion.pSetFusionContext", new
                    {
                        MessageType = MessageContext
                    }, 
                    commandType: CommandType.StoredProcedure);
                }

                int id = (int)c.Query<decimal>(@"
  insert ServiceUsers (Forename, Surname) values (@Forename, @Surname);
  select @@identity",
                    new
                    {
                        Forename = forename,
                        Surname = surname
                    }

                    ).First();

                return id;
            }            
        }

        public ServiceUser ReadServiceUser(int id)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                ServiceUser su = c.Query<ServiceUser>(@"
  select Forename, Surname from ServiceUsers where ServiceUserId = @ServiceUserId",
                    new
                    {
                        ServiceUserId = id
                    }

                    ).FirstOrDefault();

                return su;
            }
        }

    }
}
