using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Connector1.ProgressInterface;
using StructureMap.Configuration.DSL;
using System.Configuration;
using ProgressConnector.ProgressInterface;

namespace ProgressConnector.Registries
{
    public class ProgressInterfaceRegistry : Registry
    {
        public ProgressInterfaceRegistry()
        {
            string urlString = ConfigurationManager.AppSettings["urlString"];
            string userId = ConfigurationManager.AppSettings["userId"]; 
            string password = ConfigurationManager.AppSettings["password"];
            string appServerInfo = ConfigurationManager.AppSettings["appServerInfo"];
            string caller = ConfigurationManager.AppSettings["Name"];

            ProgressConnectionInfo connectionInfo = new ProgressConnectionInfo
            {
                UrlString = urlString,
                UserId = userId,
                Password = password,
                AppServerInfo = appServerInfo,
                Caller = caller
            };
            
            For<IOpenExchangeFusionMessageConvertor>().Use<OpenExchangeFusionMessageConvertor>();
            For<IOpenExchangeMessageDecoder>().Use<OpenExchangeMessageDecoder>();
            For<IReceiveMessageFromProgress>().Use<ReceiveMessageFromProgress>().Ctor<ProgressConnectionInfo>("connectionInfo").Is(connectionInfo);
            For<ISendMessageToProgress>().Use<SendMessageToProgress>().Ctor<ProgressConnectionInfo>("connectionInfo").Is(connectionInfo);
        }
    }
}
