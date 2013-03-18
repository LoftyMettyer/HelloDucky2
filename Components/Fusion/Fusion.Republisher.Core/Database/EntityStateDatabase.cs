namespace Fusion.Republisher.Core.Database
{
    using Dapper;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.SqlClient;
    using System.Linq;
    
    public class EntityStateDatabase : Fusion.Republisher.Core.Database.IEntityStateDatabase 
    {
        /// <param name="connectionString"> The connection string for the database. </param>
        public EntityStateDatabase(string connectionString)
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


        public IEnumerable<Guid> GetAllEntityRefs(string community, string messageType)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var entities = c.Query<Guid>("dbo.pEntityGetAllEntityRefs",
                    new
                    {
                        Community = community,
                        MessageType = messageType
                    },
                    commandType: CommandType.StoredProcedure);

                return entities.ToArray();
            }
        }

        public EntityState ReadEntityState(string community, string messageType, Guid entityRef)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var entities = c.Query<EntityState>("dbo.pEntityStateRead",
                    new
                    {
                        Community = community,
                        MessageType = messageType,
                        EntityRef = entityRef
                    },
                    commandType: CommandType.StoredProcedure);

                return entities.FirstOrDefault();
            }
        }


        public void UpdateEntityState(EntityState entityState)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();


                var p = new DynamicParameters();
                
                p.Add("@Community", entityState.Community);
                p.Add("@MessageType", entityState.MessageType);
                p.Add("@EntityRef", entityState.EntityRef);
                p.Add("@PrimaryEntityRef", entityState.PrimaryEntityRef);
                p.Add("@LastUpdate", entityState.LastUpdate);
                p.Add("@MessageState", entityState.MessageState);

                c.Execute("dbo.pEntityStateUpdate",
                    p,
                    commandType: CommandType.StoredProcedure);

            }
        }


        
    }
}
