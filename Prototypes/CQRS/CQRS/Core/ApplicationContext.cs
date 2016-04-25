using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;

namespace Core
{
	public class ApplicationContext : DbContext
	{
		public ApplicationContext() : base("name=Db")
		{
			Database.SetInitializer<ApplicationContext>(null);
		}


		//public  DbSet<Customer> Customers {get;set;}
		public virtual DbSet<Customer> Customers { get; set; }

		protected override void OnModelCreating(DbModelBuilder modelBuilder)
		{
			modelBuilder.Conventions.Remove<OneToManyCascadeDeleteConvention>();
			modelBuilder.Entity<Customer>().ToTable("Customer");
			base.OnModelCreating(modelBuilder);
		}
	}
}