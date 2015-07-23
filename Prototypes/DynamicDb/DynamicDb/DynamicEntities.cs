using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.ModelConfiguration.Conventions;

public partial class DynamicEntities : DbContext
{
    public DynamicEntities()
        : base("name=DynamicEntities")
    {
    }

    protected override void OnModelCreating(DbModelBuilder modelBuilder)
    {
        modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();

        //   throw new UnintentionalCodeFirstException();
    }

    public virtual DbSet<DynamicAttribute> DynamicAttributes { get; set; }
    public virtual DbSet<DynamicTemplate> DynamicTemplates { get; set; }
    public virtual DbSet<DynamicTemplateAttribute> DynamicTemplateAttributes { get; set; }
    public virtual DbSet<DynamicType> DynamicTypes { get; set; }
}