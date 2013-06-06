using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FluentNHibernate.Cfg;
using NHibernate;
using FluentNHibernate.Cfg.Db;
using NHibernate.Tool.hbm2ddl;

namespace Fusion.Publisher.SocialCare.Staff
{
    public class NHibernateSession : INHibernateSession
    {
        public NHibernateSession()
        {

        }
        private FluentConfiguration Configuration()
        {
            return Fluently.Configure()
               .Database(
                   MsSqlConfiguration
                       .MsSql2008
                       .ConnectionString(c => c
                           .Server("tcp:(local),1433")
                           .Database("NHibernateTest")
                           .Username("sa")
                           .Password("transmit")
                       ))

               .Mappings(m => m
                   .FluentMappings.AddFromAssemblyOf<Prototype.NHibernateTypeSerialization.Domain.Maps.StaffMap>()
                   .Conventions.AddFromAssemblyOf<Prototype.NHibernateTypeSerialization.Persistance.Conventions.StaffDataTypeConvention>());
        }

        private ISessionFactory CreateSessionFactory()
        {
            return Configuration().BuildSessionFactory();
        }

        private ISessionFactory sessionFactory;

        public ISession OpenSession()
        {
            if (sessionFactory == null) {
                sessionFactory = CreateSessionFactory();
            }

            return sessionFactory.OpenSession();
        }

        public void ScriptDatabase()
        {
            var config = Configuration();
            
            config.Mappings(m => {
                m.FluentMappings.ExportTo("output.hbm");
                m.AutoMappings.ExportTo("automappings.hbm");
            });

            config.ExposeConfiguration(cfg =>
                 {
                     new SchemaExport(cfg)
                         .SetOutputFile("create.sql")
                         .Create(true, true);

                 });

            config.BuildConfiguration();
        }
    }
}
