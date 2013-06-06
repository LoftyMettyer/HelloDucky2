
namespace Fusion.Test.SocialCare.Registries
{
    using System.Reflection;
    using Fusion.Core.MessageSenders;
    using StructureMap.Configuration.DSL;
    using Fusion.Core.Test;

    public class OutboundMessageWatcherRegistry : Registry
    {
        public OutboundMessageWatcherRegistry()
        {
            Scan(
                s =>
                {
                    s.Assembly(Assembly.GetExecutingAssembly());
                    s.ConnectImplementationsToTypesClosing(typeof(OutboundMessageWatcher<>));
                }

                );
        }
    }
}
