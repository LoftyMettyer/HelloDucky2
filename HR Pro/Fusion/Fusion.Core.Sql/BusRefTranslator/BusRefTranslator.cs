// --------------------------------------------------------------------------------------------------------------------
// <copyright file="BusRefTranslator.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Provide a facility to translate messages between bus and local ids
// </summary>
// --------------------------------------------------------------------------------------------------------------------


namespace Fusion.Core.Sql
{
    using System;
    using System.Data;
    using System.Data.SqlClient;
    using Dapper;

    /// <summary>
    /// Translate ids using sql stored procedures
    /// </summary>
    public class BusRefTranslator : IBusRefTranslator
    {

        /// <summary>
        /// Initializes a new instance of the BusRefTranslator class.
        /// </summary>
        /// <param name="connectionString"> The connection string for the database. </param>
        public BusRefTranslator(string connectionString)
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


        public void SetBusRef(string translationName, string localId, Guid busRef)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                c.Execute("fusion.pIdTranslateSetBusRef",
                    new
                    {
                        LocalId = localId,
                        TranslationName = translationName,
                        BusRef = busRef
                    },
                    commandType: CommandType.StoredProcedure);
            }
        }
        
        public BusTranslationResults TryGetBusRef(string translationName, string localId, bool canCreate = false)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var p = new DynamicParameters();
                p.Add("@LocalId", localId);
                p.Add("@TranslationName", translationName);
                p.Add("@CanGenerate", canCreate);
                p.Add("@BusRef", dbType: DbType.Guid, direction: ParameterDirection.Output);
                p.Add("@DidGenerate", dbType: DbType.Boolean, direction: ParameterDirection.Output);

                c.Execute("fusion.pIdTranslateGetBusRef",
                   p,
                    commandType: CommandType.StoredProcedure);

                return new BusTranslationResults
                {
                    BusRef = p.Get<Guid?>("@BusRef"),
                    LocalId = localId,
                    BusRefNewlyCreated = p.Get<bool>("@DidGenerate"),
                };
            }
        }

        public string GetLocalRef(string translationName, Guid busRef)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var p = new DynamicParameters();
                p.Add("@BusRef", busRef);
                p.Add("@TranslationName", translationName);
                p.Add("@LocalId", dbType: DbType.String, size: 25, direction: ParameterDirection.Output);

                c.Execute("fusion.pIdTranslateGetLocalId",
                   p,
                    commandType: CommandType.StoredProcedure);

                return p.Get<string>("@LocalId");
            }
        }
    }
}
