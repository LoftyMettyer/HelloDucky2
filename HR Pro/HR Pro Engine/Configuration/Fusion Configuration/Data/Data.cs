using System.Linq;
using NHibernate;
using NHibernate.Dialect;
using NHibernate.Cfg;
using NHibernate.Cfg.MappingSchema;
using ConfOrm;
using ConfOrm.Patterns;
using ConfOrm.NH;
using ConfOrm.Shop.CoolNaming;

namespace Fusion
{
	public class Data
	{
		public static ISessionFactory SessionFactory { get; set; }

		public static ISessionFactory BuildSessionFactory(string connectionString)
		{
			var config = new Configuration();

			config.DataBaseIntegration(db => {
													db.Dialect<MsSql2008Dialect>();		
													db.ConnectionString = connectionString;
			                           	db.SchemaAction = SchemaAutoAction.Validate;
			                           });

			config.AddDeserializedMapping(GetMappings(), "");

			return config.BuildSessionFactory();
		}

		private static HbmMapping GetMappings()
		{
			var entities = typeof (Entity).Assembly.GetTypes().Where(t => t.BaseType == typeof (Entity)).ToList();

			var orm = new ObjectRelationalMapper();
			orm.Patterns.PoidStrategies.Add(new IdentityPoidPattern());
			orm.TablePerClass(entities);

			var mapper = new Mapper(orm, new CoolPatternsAppliersHolder(orm));

			mapper.AddPropertyPattern(m => m.GetPropertyOrFieldType() == typeof(string), a => a.Type(NHibernateUtil.AnsiString));
			
			mapper.Class<Entity>(ca => ca.DynamicUpdate(true));

			mapper.Class<Table>(ca => ca.Subselect("SELECT tableid as id, tablename as name FROM ASRSysTables"));

			mapper.Class<Column>(ca => ca.Subselect("SELECT columnid as ID, columnname as Name, tableID, datatype, size, decimals, lookupTableId FROM ASRSysColumns"));

			mapper.Class<FusionCategory>(ca => { ca.Schema("fusion"); ca.Table("Category"); });

			mapper.Class<FusionElement>(ca => { ca.Schema("fusion"); ca.Table("Element"); });

			mapper.Class<FusionLog>(ca => { ca.Schema("fusion"); ca.Table("MessageTracking"); });

			mapper.Class<FusionMessage>(ca => { ca.Schema("fusion"); ca.Table("Message"); });

			mapper.Class<FusionMessageElement>(ca => { ca.Schema("fusion"); ca.Table("MessageElements"); });

			var mapping = mapper.CompileMappingFor(entities);

			return mapping;
		}
	}
}