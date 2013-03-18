using Fusion.Republisher.Core.Database;
using StructureMap.Configuration.DSL;

namespace Fusion.Republisher.Core.Registries
{
    public class DatabaseAccessRegistry : Registry
    {
        public DatabaseAccessRegistry()
        {
            For<IEntityStateDatabase>().Use<EntityStateDatabase>().Ctor<string>("connectionString").EqualToAppSetting("connectionString");                  
        }
    }
}
