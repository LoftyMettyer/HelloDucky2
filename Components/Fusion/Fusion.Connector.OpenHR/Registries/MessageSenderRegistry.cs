using StructureMap.Configuration.DSL;
using System.Reflection;
using Fusion.Core.MessageSenders;
using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Registries
{
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
