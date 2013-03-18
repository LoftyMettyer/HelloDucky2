namespace ProgressConnector.MessageHandlers
{
    using System;
    using Fusion.Core;
    using Fusion.Messages.General;
    using log4net;
    using NServiceBus;
    using ProgressConnector.ProgressInterface;
    using StructureMap.Attributes;

    public class StandardMessageHandler : IHandleMessages<FusionMessage>
    {
         [SetterProperty]
        public ISendMessageToProgress ProgressInterface
        {
            get;
            set;
        }

        //[SetterProperty]
        //public IOpenExchangeFusionMessageConvertor MessageConvertor
        //{
        //    get;
        //    set;
        //}

        private static readonly ILog Logger = LogManager.GetLogger(typeof(StandardMessageHandler));

        public void Handle(FusionMessage message) {

            Logger.Info(string.Format("Connector received " + message.GetMessageName() + " with Id {0} from {1} - xml {2}.", message.Id, message.Originator, message.Xml));

            Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));
      
            //string openExchangeMessage = MessageConvertor.BuildOpenExchangeMessage(message);

            var sendResults = ProgressInterface.SendMessage(message);
            if (sendResults.Error)
            {
                throw new ApplicationException("Progress interface reports " + sendResults.ErrorText);
            }

        }

    }
}
