

namespace Fusion.LogService.Registries
{
    using StructureMap.Configuration.DSL;
    using Fusion.LogService.DatabaseAccess;
    using System.Configuration;
    using System;

    public class DatabaseAccessRegistry : Registry
    {
        public DatabaseAccessRegistry()
        {
            string connectionString = ConfigurationManager.AppSettings["connectionString"];

            if (String.IsNullOrWhiteSpace(connectionString))
            {
                For<ILogDatabase>().Use<NullLogDatabase>();
            }
            else
            {
                For<ILogDatabase>().Use<LogDatabase>().Ctor<string>("connectionString").Is(connectionString);

            }
        } 
    }
}
