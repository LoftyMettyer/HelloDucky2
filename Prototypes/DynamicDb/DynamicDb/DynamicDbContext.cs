using System;
using System.Collections.Generic;
using System.Data.Entity;
using MagicDbModelBuilder;
using System.Linq;
using System.Diagnostics;

public partial class DynamicDbContext : DbContext
{
    public DynamicDbContext()
        : base("name=DynamicDbContext")
    {
        Database.SetInitializer(new NullDatabaseInitializer<DynamicDbContext>());
    }

    public void AddTable(Type type)
    {
        _tables.Add(type.Name, type);
    }

    private Dictionary<string, Type> _tables = new Dictionary<string, Type>();

    protected override void OnModelCreating(DbModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);
        var entityMethod = modelBuilder.GetType().GetMethod("Entity");

        foreach (var table in _tables)
        {
            entityMethod.MakeGenericMethod(table.Value).Invoke(modelBuilder, new object[] { });
            foreach (var pi in (table.Value).GetProperties())
            {
                if (pi.Name == "Id")
                    modelBuilder.Entity(table.Value).HasKey(typeof(int), "Id");
                else
                    modelBuilder.Entity(table.Value).StringProperty(pi.Name);
            }
        }
    }

    //public Type GetTable(string name)
    //{
    //    //return _tables.First();
    //}

    //public virtual DbSet<DynamicAttribute> DynamicAttributes { get; set; }

    public Type GetTable(string name)
    {
        return _tables[name];
    }

    //public dbSet

}