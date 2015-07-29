using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Data.SqlClient;
using System.Linq;
using OpenHRNexus.Interfaces.Common;
using OpenHRNexus.Repository.DatabaseClasses;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.Messages;

namespace OpenHRNexus.Repository.SQLServer {
	public class SqlAuthenticateRepository : DbContext, IWelcomeMessageDataRepository, IAuthenticateRepository {

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

		public IEnumerable<string> GetUserPermissions(Guid userId)
		{
		//	var result = Roles.Select(c => new Role { c.id, c.Name });

		//	var result = Roles.Select(c => new { c.Name, c.id});

		//	return Roles.FirstOrDefault(m => m.Id == userId).Name;

			//(
			//var myResult = Roles.Where(c => c.Name == "someName").Select(c => c.Name);
			var result = UserRoles
					.Where(c => c.UserID == userId)
					.Select(c => c.Name);

			return result.ToList();
			//throw new NotImplementedException();


		}

		public WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language)
		{
			var userIDParameter = userID.HasValue ?
					new SqlParameter("UserId", userID) :
					new SqlParameter("UserId", typeof(Guid));

			var languageParameter = language != null ?
					new SqlParameter("Language", language) :
					new SqlParameter("Language", typeof(string));

			var result = Database
							.SqlQuery<WelcomeDataMessage>("GetWelcomeMessageData @UserId, @Language", userIDParameter, languageParameter);

			var message = result.FirstOrDefault();

			return message;

		}

		public virtual DbSet<Role> Roles { get; set; }
		public virtual DbSet<UserRole> UserRoles { get; set; }

	//	public virtual DbSet<INexusUser> Users { get; set; }

		//public IEnumerable<string> GetUserRoles(Guid userId)
		//{
		//	var roles = new List<string> {"Admin", "Default User"};
		//	return roles;
		//}

		//public IEnumerable<string> GetUserClaims(Guid userId)
		//{
		//	var roles = new List<string> { "Admin", "Default User" };
		//	return roles;
		//}

	}
}