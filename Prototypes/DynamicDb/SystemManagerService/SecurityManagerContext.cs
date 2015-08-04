using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;

namespace SystemManagerService
{
    public class SecurityManagerContext : DbContext
    {
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
        }

    }
}
