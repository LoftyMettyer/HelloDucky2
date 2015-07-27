using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Data.SqlClient;
using System.Linq;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.Messages;

namespace OpenHRNexus.Repository.SQLServer {
	public class SqlAuthenticateRepository : DbContext, IAuthenticateRepository {

		public SqlAuthenticateRepository()
			: base("name=SqlAuthenticateRepository")
		{
		}

		protected override void OnModelCreating(DbModelBuilder modelBuilder)
		{
			modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
		}		

		public RegisterNewUserMessage RequestAccount(string email)
		{

			var emailParameter = email != null ?
					new SqlParameter("email", email) :
					new SqlParameter("email", typeof(string));

			var result = Database
					.SqlQuery<RegisterNewUserMessage>("RegisterNewUser @email", emailParameter);

			var message = result.FirstOrDefault();

			return message;

		}
	}
}