using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;

namespace Nexus.Repository
{
    public class SQLRepositoryContext : DbContext
    {
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
        }

    }
}
