using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.Repositories.MySQL {
	public class MySQLtbuser_LanguagesRepository : Itbuser_LanguagesRepository {
		public List<tbuser_Languages> List() {
			List<tbuser_Languages> _listingList = new List<tbuser_Languages>();

			const string cs = @"Server=Localhost;Uid=user;Pwd=password;Database=database"; //MySQL connection string

			MySqlConnection conn = null;
			MySqlDataReader rdr = null;

			conn = new MySqlConnection(cs);
			conn.Open();

			string stm = "SELECT * FROM tbuser_languages";
			MySqlCommand cmd = new MySqlCommand(stm, conn);
			rdr = cmd.ExecuteReader();

			while (rdr.Read()) {
				_listingList.Add(
					new tbuser_Languages {
						ID = rdr.GetInt32(0),
						ID_1 = rdr.GetInt32(1),
						Language_Level=rdr.GetDecimal(2),
						Language_Name= rdr.GetString(3),
						Spoken_Fluency = rdr.GetString(4),
						Written_Fluency = rdr.GetString(3),
					});
			}

			return _listingList;
		}
	}
}
