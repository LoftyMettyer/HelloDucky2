using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using StructureMap.Configuration.DSL;
using System.Reflection;
using Fusion.Core.MessageSenders;

namespace ExampleConnector.Registries
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
