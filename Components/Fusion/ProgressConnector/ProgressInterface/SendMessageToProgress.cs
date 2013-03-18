
namespace ProgressConnector.ProgressInterface
{
    using Fusion.Core;
    using Fusion.Messages.General;
    using oxProcess;

    public class SendMessageToProgress : ISendMessageToProgress
    {
        public SendMessageToProgress(ProgressConnectionInfo connectionInfo)
        {
            this.urlString = connectionInfo.UrlString;
            this.userId = connectionInfo.UserId;
            this.password = connectionInfo.Password;
            this.appServerInfo = connectionInfo.AppServerInfo;
            this.caller = connectionInfo.Caller;
        }

        private string urlString;
        private string userId;
        private string password;
        private string appServerInfo;
        private string caller;

        public ProgressSendStatus SendMessage(FusionMessage message)
        {

            using (oxProcessListener oxProcess = new oxProcessListener(urlString, userId, password, appServerInfo))
            {
                bool errorFlag;
                string errorString;

                string result = oxProcess.connectortoprogress(
                    message.Xml,
                    caller, 
                    message.Id.ToString(),
                    message.Originator,
                    message.CreatedUtc,
                    message.SchemaVersion,
                    message.Community,
                    message.EntityRef.HasValue ? message.EntityRef.Value.ToString() : null,
                    message.GetMessageName(),
                    message.PrimaryEntityRef.HasValue ? message.PrimaryEntityRef.Value.ToString() : null,
                    out errorFlag, 
                    out errorString);

                return new ProgressSendStatus
                {
                    Error = errorFlag,
                    ErrorText = errorString
                };
            }         

        }
    }
}
