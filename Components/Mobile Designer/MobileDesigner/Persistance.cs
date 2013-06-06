using System.Linq;
using NHibernate;
using NHibernate.Cfg;
using NHibernate.Cfg.MappingSchema;
using ConfOrm;
using ConfOrm.Patterns;
using ConfOrm.NH;
using ConfOrm.Shop.CoolNaming;
using NHibernate.Mapping.ByCode;
using TypeExtensions = ConfOrm.TypeExtensions;

namespace MobileDesigner
{
    public class DataManager
    {
        public static ISessionFactory BuildSessionFactory(string connectionString)
        {
            var config = new Configuration();

            config.DataBaseIntegration(db =>
            {
                db.Dialect<NHibernate.JetDriver.JetDialect>();
                db.Driver<JetDriverFixed>();
                db.ConnectionString = connectionString;
            });

            config.AddDeserializedMapping(GetMappings(), "");

            return config.BuildSessionFactory();
        }

        private static HbmMapping GetMappings()
        {
            var entities = typeof(Entity).Assembly.GetTypes().Where(t => t.BaseType == typeof(Entity)).ToList();

            var orm = new ObjectRelationalMapper();
            orm.Patterns.PoidStrategies.Add(new IdentityPoidPattern());
            orm.TablePerClass(entities);
            orm.ManyToMany<UserGroup, Workflow>();

            var mapper = new Mapper(orm, new CoolPatternsAppliersHolder(orm));

            mapper.AddPropertyPattern(m => TypeExtensions.GetPropertyOrFieldType(m) == typeof(string), a => a.Type(NHibernateUtil.AnsiString));

            mapper.Class<Entity>(ca => ca.DynamicUpdate(true));

            mapper.Class<Layout>(ca => ca.Table("tmpmobileformlayout"));

            mapper.Class<Element>(ca => ca.Table("tmpmobileformelements"));

            mapper.Class<Picture>(ca =>
            {
                ca.Table("tmpPictures");
                ca.Id(p => p.Id, m =>
                                     {
                                         m.Column("PictureID"); 
                                         m.Generator(Generators.Assigned);
                                     } );
                ca.Property(p => p.Image, m => m.Column("Picture"));
                ca.Property(p => p.Type, m => m.Column("PictureType"));
            });

            mapper.Class<Workflow>(ca => ca.Table("tmpworkflows"));

            mapper.Class<UserGroup>(ca =>
                                        {
                                            ca.Table("tmpGroups");
                                            ca.List(b => b.MobileWorkflows, cm =>
                                                                                {
                                                                                    cm.Table("tmpmobilegroupworkflows");
                                                                                    cm.Index(lim => lim.Column("Pos"));
                                                                                });
                                        });

            var mapping = mapper.CompileMappingFor(entities);

            return mapping;
        }
    }

    public class JetDriverFixed : NHibernate.JetDriver.JetDriver
    {
        public override bool UseNamedPrefixInSql { get { return true; } }
    }
}
