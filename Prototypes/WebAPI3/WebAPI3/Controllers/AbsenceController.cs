using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using WebAPI3.Models;
using Dapper;

namespace WebAPI3.Controllers
{
    public class AbsenceController : ApiController
    {
        static string connStr = ConfigurationManager.ConnectionStrings["OpenHR"].ConnectionString;
        SqlConnection con = new SqlConnection(connStr);

        //public AbsenceController() { }

        //public IEnumerable<Absence> GetAll(){
        //    return null;
        //}

        //public IEnumerable<Absence> GetAll(int PersonID)
        //{

        //    var absences = con.Query<Absence>(@"SELECT * FROM Absence WHERE ID_1 = @id", new { ID = PersonID });

        //    return absences;
        //}

        /// <summary>
        /// Gets absence
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public IEnumerable<Absence> GetAll(int id)
//public IHttpActionResult GetAll(int PersonID)
        {
            con.Open();
  //          var abs = con.Query<Absence>(@"SELECT * FROM Absence WHERE ID_1 = @id", new { ID = 1 }).FirstOrDefault();
//            return Ok(abs);

            var absences = con.Query<Absence>(@"SELECT * FROM Absence WHERE ID_1 = @id", new { ID = id });

            return absences;

        }

        public void ExecuteSomething()
        {
            var i = 1;


        }

    }
}
