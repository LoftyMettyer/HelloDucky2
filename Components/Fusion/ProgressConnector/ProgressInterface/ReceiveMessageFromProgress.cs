using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using oxProcess;

namespace ProgressConnector.ProgressInterface
{
    public class ReceiveMessageFromProgress : IReceiveMessageFromProgress
    {
        public ReceiveMessageFromProgress(ProgressConnectionInfo connectionInfo)
        {
            this.urlString = connectionInfo.UrlString;
            this.userId = connectionInfo.UserId;
            this.password = connectionInfo.Password;
            this.appServerInfo = connectionInfo.AppServerInfo;
        }

        private string urlString;
        private string userId;
        private string password;
        private string appServerInfo;

        public RawOpenExchangeData ReceiveOneMessage()
        {
             using (oxProcessListener oxProcess = new oxProcessListener(urlString, userId, password, appServerInfo)) {
                 string content;
                 bool errorFlag;
                 string errorString;
                 oxProcess.progresstoconnector(out content, out errorFlag, out errorString);
                 if (errorFlag == true)
                 {
                     throw new ApplicationException("OpenExchange error " + errorString);
                 }

                 if (String.IsNullOrEmpty(content))
                     return null;

                 return new RawOpenExchangeData
                 {
                     MessageRequestXml = content
                 };
            }
       }


        public void AcknowledgeSent(Guid id)
        {
            using (oxProcessListener oxProcess = new oxProcessListener(urlString, userId, password, appServerInfo))
            
            {
                bool errorFlag;
                string errorString;
 
                oxProcess.connectorconfirmsmessageonbus(id.ToString(), out errorFlag, out errorString);
                if (errorFlag == true)
                {
                    throw new ApplicationException("Progress Interface error " + errorString);
                }
            }
 
        }
    }
}
