using System;
using System.Threading;
using Fusion.Core.MessageSenders;
using Fusion.Core.OutboundFilters;
using log4net;
using NServiceBus;
using ProgressConnector.BusTypeBuilder;
using ProgressConnector.ProgressInterface;
using StructureMap.Attributes;
using Fusion.Messages.General;

namespace ProgressConnector
{
    public class OpenExchangeMessageReceiver : IWantToRunAtStartup
    {
        public IBus Bus
        {
            get;
            set;
        }

        [SetterProperty]
        public IBusTypeBuilder OutboundBuilder
        {
            get;
            set;
        }

        [SetterProperty]
        public IReceiveMessageFromProgress OpenExchangeListener
        {
            get;
            set;
        }

        [SetterProperty]
        public IOpenExchangeMessageDecoder MessageDecoder
        {
            get;
            set;
        }


        private static readonly ILog Logger = LogManager.GetLogger(typeof(OpenExchangeMessageReceiver));

        public void Run()
        {

            Logger.Info("Polling OpenExchange...");


            for (; ; ) {
            
                RawOpenExchangeData rawOpenExchangeMessage = null;
                try
                {
                    rawOpenExchangeMessage = OpenExchangeListener.ReceiveOneMessage();
                }
                catch (Exception e) 
                {
                    Logger.Error("Error reported from OpenExchange", e);
                    Thread.Sleep(TimeSpan.FromSeconds(30));

                    continue;
                }

                if (rawOpenExchangeMessage == null)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(10));
                } else if (rawOpenExchangeMessage != null)
                {
                    Logger.Info("Received command from progress queue: " + rawOpenExchangeMessage.MessageRequestXml);
                    
                    var openExchangeMessage = MessageDecoder.Decode(rawOpenExchangeMessage);

                    if (openExchangeMessage is OpenExchangeMessage)
                    {
                        SendMessage((OpenExchangeMessage)openExchangeMessage);
                    }
                    else if (openExchangeMessage is OpenExchangeLogMessage)
                    {
                        LogMessage((OpenExchangeLogMessage)openExchangeMessage);
                    }
                    else if (openExchangeMessage is OpenExchangeIdTranslation)
                    {
                        LogIdTranslation((OpenExchangeIdTranslation)openExchangeMessage);
                    }
                    else
                    {
                        Logger.Error("Message unrecognised! " + rawOpenExchangeMessage.MessageRequestXml);
                    }
                }
            }
        }

        private bool LogMessage(OpenExchangeLogMessage openExchangeMessage)
        {
            Logger.InfoFormat("Log request received {0}/{1}", openExchangeMessage.Message, openExchangeMessage.Id);

            Bus.Send(
                  new LogMessage
                  {
                      Id = openExchangeMessage.Id,
                      Message = openExchangeMessage.Message,
                      Community = openExchangeMessage.Community,
                      MessageDescription = openExchangeMessage.MessageDescription,
                      MessageId = openExchangeMessage.MessageId,
                      Source = openExchangeMessage.Source,
                      TimeUtc = openExchangeMessage.TimeUtc,
                      EntityRef = openExchangeMessage.EntityRef,
                      LogLevel = (FusionLogLevel)Enum.Parse(typeof(FusionLogLevel), openExchangeMessage.LogLevel)                      
                  });

            OpenExchangeListener.AcknowledgeSent(openExchangeMessage.Id);
            
            return true;
        }

        private bool LogIdTranslation(OpenExchangeIdTranslation openExchangeMessage)
        {
            Logger.InfoFormat("Log id translation received {0}/{1} <--> {2}", openExchangeMessage.EntityName, openExchangeMessage.LocalId, openExchangeMessage.BusRef);

            Bus.Send(
                  new LogTranslationMessage
                  {
                      Id = openExchangeMessage.Id,
                      Community = openExchangeMessage.Community,
                      MessageId = openExchangeMessage.MessageId,
                      Source = openExchangeMessage.Source,
                      TimeUtc = openExchangeMessage.TimeUtc,
                      BusRef = openExchangeMessage.BusRef,
                      LocalId = openExchangeMessage.LocalId,
                      EntityName = openExchangeMessage.EntityName                      
                  });

            OpenExchangeListener.AcknowledgeSent(openExchangeMessage.Id);

            return true;
        }

        private bool SendMessage(OpenExchangeMessage openExchangeMessage)
        {
            Logger.InfoFormat("Message received {0}/{1}", openExchangeMessage.MessageType, openExchangeMessage.Id);

            var fusionMessage = OutboundBuilder.Build(openExchangeMessage);

            if (fusionMessage != null)
            {
                if (fusionMessage is ICommand)
                {
                    Bus.Send(fusionMessage);
                }
                else if (fusionMessage is IEvent)
                {
                    Bus.Publish(fusionMessage);
                }
                else
                {
                    Logger.ErrorFormat("Message type {0} is not an ICommand or an IEvent", fusionMessage.GetType());
                }

                OpenExchangeListener.AcknowledgeSent(fusionMessage.Id);
            }
            return true;
        }


        public void Stop()
        {

        }
    }
}