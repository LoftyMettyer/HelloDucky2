using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;

namespace Nexus.Sql_Repository {
	public class SqlRepositoryContext : DbContext {
		protected override void OnModelCreating(DbModelBuilder modelBuilder) {
			modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();

  //          modelBuilder.Entity<>

		}

	}
}
