using System;
using System.Collections.Generic;
using System.Data.Services;
using System.Data.Services.Common;
using System.Data.Services.Providers;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace DynamicDb
{
    [ServiceBehavior(IncludeExceptionDetailInFaults = true)]
    public class DynamicODataService : EntityFrameworkDataService<DynamicDbContext>
    {

        public static void InitializeService(DataServiceConfiguration config)
        {
            config.SetEntitySetAccessRule("*", EntitySetRights.All);
            config.SetServiceActionAccessRule("*", ServiceActionRights.Invoke);
            config.DataServiceBehavior.AcceptProjectionRequests = true;
            config.UseVerboseErrors = true;
            config.DataServiceBehavior.MaxProtocolVersion = DataServiceProtocolVersion.V3;
        }

        protected override DynamicDbContext CreateDataSource()
        {
            var result = base.CreateDataSource();
            var dcf = new DynamicClassFactory();

            var context = new DynamicEntities();
            var templates = (from t in context.DynamicTemplates.Include("DynamicTemplateAttributes").Include("DynamicTemplateAttributes.DynamicAttribute")
                             select t);

            foreach (var dynamicTemplate in templates)
            {
                var type = CreateType(dcf, dynamicTemplate.Name, dynamicTemplate.DynamicTemplateAttributes);
                result.AddTable(type);
            }

            return result;
        }

        private Type CreateType(DynamicClassFactory dcf, string name, ICollection<DynamicTemplateAttribute> dynamicAttributes)
        {
            var props = dynamicAttributes.ToDictionary(da => da.DynamicAttribute.Name, da => typeof(string));
            var t = dcf.CreateDynamicType<BaseDynamicEntity>(name, props);
            return t;
        }

        public void CreateTable(string template)
        {
            var context = new DynamicEntities();
            var qry = (from dt in context.DynamicTemplates.Include("DynamicTemplateAttributes")
                .Include("DynamicTemplateAttributes.DynamicAttribute")
                       select dt).FirstOrDefault(dt => dt.Name == template);

         

            if (qry == null)
                throw new ArgumentException(string.Format("The template {0} does not exist", template));
            var ct = new StringBuilder();
            ct.AppendFormat("CREATE TABLE {0} (Id int IDENTITY(1,1) NOT NULL, ", qry.Name);
            foreach (var dta in qry.DynamicTemplateAttributes)
            {
                ct.AppendFormat("{0} nvarchar(255) NULL, ", dta.DynamicAttribute.Name);
            }
            ct.AppendFormat("CONSTRAINT [PK_{0}] PRIMARY KEY CLUSTERED", qry.Name);
            ct.AppendFormat("(Id ASC) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF," +
                            "ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]");
            var ts = ct.ToString();
            context.Database.ExecuteSqlCommand(ts);
        }

        public DynamicDbContext GetDataSource()
        {
            var dataSource = CreateDataSource();

            //dataSource.ta

            var blah2 = new DynamicDbContext();
            //blah2.DoSomethingToTripTheOnModelCreating();
            //  blah2

            //var blah3 = dataSource.

            //return dataSource._tables["SomeTable"];
            var blah3 = dataSource.GetTable("SomeTable");


            //var qry = (from dt in context.DynamicTemplates.Include("DynamicTemplateAttributes").Include("DynamicTemplateAttributes.DynamicAttribute")

            //var qry = (from dt in blah2.GetTable("SomeTable") select dt).FirstOrDefault(dt => dt.Name == template);


            return dataSource;


    //        JsonConvert.SerializeObject(entity, Formatting.Indented, settings);

            //  return Json("hello", JsonRequestBehavior.AllowGet)

            //     return dataSource;
        }

    }
}
