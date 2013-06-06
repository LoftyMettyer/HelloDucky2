
namespace Fusion.Test.SocialCare.Registries
{
    using System.Reflection;
    using Fusion.Core.MessageSenders;
    using StructureMap.Configuration.DSL;

    public class MessageSenderRegistry : Registry
    {
        public MessageSenderRegistry()
        {
            For<IMessageSenderInvoker>().Use<MessageSenderInvoker>();

            Scan(
                s =>
                {
                    s.Assembly(Assembly.GetExecutingAssembly());
                    s.ConnectImplementationsToTypesClosing(typeof(IMessageSender<>));
                }

                );
        }
    }
}
