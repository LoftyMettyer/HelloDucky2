using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Republisher.SocialCare.Domain.Entities;
using System.Data.SqlClient;

using Dapper;
using System.Data;

namespace Fusion.Publisher.SocialCare.Database
{
    public class StaffRepository : IStaffRepository
    {
            /// <param name="connectionString"> The connection string for the database. </param>
        public StaffRepository(string connectionString)
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


        public IEnumerable<Staff> GetStaffMember(Guid staffRef)
        {
            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var staff = c.Query<Staff>("dbo.pGetStaff",
                    new
                    {
                        StaffRef = staffRef,                       
                    },
                    commandType: CommandType.StoredProcedure);

                return staff.ToArray();
            }
        }

        public void UpdateStaffMember(Guid staffRef, IEnumerable<Staff> staffList)
        {
            var staffArray = staffList.ToArray();

            if (staffArray.Any(x => x.StaffRef != staffRef))
            {
                throw new ArgumentException("List contains a record which does not belong", "staffList");
            }

            var overlappingList =
                from a in staffList
                from b in staffList
                where !Object.Equals(a, b) &&
                a.EffectiveFrom <= b.EffectiveTo && a.EffectiveTo >= b.EffectiveFrom
                select new
                {
                    AFrom = a.EffectiveFrom, 
                    ATo = a.EffectiveTo, 
                    BFrom = b.EffectiveFrom, 
                    BTo = b.EffectiveTo
                };

            if (overlappingList.FirstOrDefault() != null)
            {
                throw new ArgumentException("List contains overlapping records", "staffList");
            }

            using (SqlConnection c = new SqlConnection(ConnectionString))
            {
                c.Open();

                //var staff = c.Query<Staff>("dbo.pGetStaff",
                //    new
                //    {
                //        StaffRef = staffRef,
                //    },
                //    commandType: CommandType.StoredProcedure);
               
            }
        }
        

        //public BusTranslationResults TryGetBusRef(string translationName, string localId, bool canCreate = false)
        //{
        //    using (SqlConnection c = new SqlConnection(ConnectionString))
        //    {
        //        c.Open();

        //        var p = new DynamicParameters();
        //        p.Add("@LocalId", localId);
        //        p.Add("@TranslationName", translationName);
        //        p.Add("@CanGenerate", canCreate);
        //        p.Add("@BusRef", dbType: DbType.Guid, direction: ParameterDirection.Output);
        //        p.Add("@DidGenerate", dbType: DbType.Boolean, direction: ParameterDirection.Output);

        //        c.Execute("fusion.pIdTranslateGetBusRef",
        //           p,
        //            commandType: CommandType.StoredProcedure);

        //        return new BusTranslationResults
        //        {
        //            BusRef = p.Get<Guid?>("@BusRef"),
        //            LocalId = localId,
        //            BusRefNewlyCreated = p.Get<bool>("@DidGenerate"),
        //        };
        //    }
        //}


    }
}
