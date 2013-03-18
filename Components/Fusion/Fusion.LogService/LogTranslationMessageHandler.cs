using log4net;
using NServiceBus;
using Fusion.Messages.General;
using Fusion.LogService.DatabaseAccess;
using StructureMap.Attributes;

namespace Fusion.LogService
{
    public class LogTranslationMessageHandler : IHandleMessages<LogTranslationMessage>
    {
        [SetterProperty]
        public ILogDatabase LogDatabase
        {
            get;
            set;
        }

        public void Handle(LogTranslationMessage message)
        {
            Logger.Info(string.Format("Log service received Translation message with Id {0} time {1} translate from {2}:{3} to {4}", message.Id, message.TimeUtc, message.EntityName, message.LocalId, message.BusRef));

            LogDatabase.AddTranslationRecord(message.TimeUtc, message.MessageId, message.Source, message.Community, message.EntityName, message.BusRef, message.LocalId);
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(LogTranslationMessageHandler));
    }
}
