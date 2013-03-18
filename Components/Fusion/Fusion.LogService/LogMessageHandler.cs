using log4net;
using NServiceBus;
using Fusion.Messages.General;
using Fusion.LogService.DatabaseAccess;
using StructureMap.Attributes;

namespace Fusion.LogService
{
    public class LogMessageHandler : IHandleMessages<LogMessage>
    {

        [SetterProperty]
        public ILogDatabase LogDatabase
        {
            get;
            set;
        }

        public void Handle(LogMessage message)
        {
            Logger.Info(string.Format("Log service received EventMessage with Id {0} time {1} wrt MessageRef {2} - message {3}.", message.Id,
                message.TimeUtc, message.MessageId.HasValue ? message.MessageId.Value.ToString() : "--", message.Message));

            char logLevel = ' ';
            switch (message.LogLevel)
            {
                case FusionLogLevel.Info:
                    logLevel = 'I'; break;
                case FusionLogLevel.Error:
                    logLevel = 'E'; break;
                case FusionLogLevel.Warning:
                    logLevel = 'W'; break;
                case FusionLogLevel.Fatal:
                    logLevel = 'F'; break;
            }

            LogDatabase.AddLogEntry(
                message.Id, message.MessageName, message.MessageId, message.Community, message.Source, message.EntityRef, message.PrimaryEntityRef, message.TimeUtc, logLevel, message.Message);
       }

        private static readonly ILog Logger = LogManager.GetLogger(typeof (LogMessageHandler));
    }
}
