using Fusion.Core;
using Fusion.Core.Logging;
using Fusion.Messages.General;
using Fusion.Republisher.Core;
using Fusion.Republisher.Core.Configuration;
using Fusion.Republisher.Core.Database;
using Fusion.Republisher.Core.MessageProcessors;
using Fusion.Republisher.Core.MessageStateSerializer;
using log4net;
using NServiceBus;
using StructureMap.Attributes;
using System;

namespace Fusion.Republisher.Core
{
    public class StateStoreMessageRepublisher <Source, Destination, MessageDefinition> 
        : IHandleMessages<Source>

        where Source : FusionMessage, ICommand                                                                     
        where Destination : FusionMessage, IEvent, new()
        where MessageDefinition : IMessageDefinition, new()
    {
        /// <summary>
        /// Injected reference to NServiceBus
        /// </summary>
        public IBus Bus
        {
            get;
            set;
        }

        [SetterProperty]
        public IEntityStateDatabase StateDatabase
        {
            get;
            set;
        }

        [SetterProperty]
        public IMessageStateSerializer StateSerializer
        {
            get;
            set;
        }

        [SetterProperty]
        public IFusionLogService FusionLogger
        {
            get;
            set;
        }

        [SetterProperty]
        public IFusionMessageProcessor MessageProcessor
        {
            get;
            set;
        }

        [SetterProperty]
        public IFusionConfiguration Configuration
        {
            get;
            set;

        }

        private readonly ILog Logger;

        public StateStoreMessageRepublisher()
        {
            this.Logger = LogManager.GetLogger(this.GetType());
        }

        public void Handle(Source message)
        {
            this.Logger.Info(string.Format("Fusion Performing State republish received {0} from {1} with Id {2}.", message.GetMessageName(), message.Originator, message.Id));
            this.Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));

            this.FusionLogger.LogMessageReceived(message);

            string outputXml;

            string myCommunity = Configuration.Community;
            if (message.Community != myCommunity)
            {
                this.Logger.Info(String.Format("Message received with unexpected community - expected {0}", myCommunity));

                FusionLogger.InfoMessageNonTransactional(message, Messages.General.FusionLogLevel.Error,
                    String.Format("Message received with unexpected community - expected {0}", myCommunity));

                throw new InvalidOperationException();
            }

            if (this.Configuration.StoreState == true)
            {

                var currentState = this.StateDatabase.ReadEntityState(message.Community, message.GetMessageName(), message.EntityRef.Value);

                if (currentState == null)
                {
                    currentState = new EntityState
                    {
                        Community = message.Community,
                        EntityRef = message.EntityRef.Value,
                        PrimaryEntityRef = message.PrimaryEntityRef.Value,
                        MessageType = message.GetMessageName(),
                    };
                }

                if (message.CreatedUtc < currentState.LastUpdate)
                {
                    this.FusionLogger.InfoMessageNonTransactional(message, Messages.General.FusionLogLevel.Warning,
                        String.Format("Message received with created date {0} before current latest {1}", message.CreatedUtc, currentState.LastUpdate));

                    return;
                }

                MessagePersistedState messageData = this.StateSerializer.Deserialize(currentState.MessageState);

                // Should be retrieved somehow?
                IMessageDefinition messageDefinition = (new MessageDefinition());

                // Update state

                this.MessageProcessor.UpdateStateFromMessage(messageDefinition, messageData, message.Xml);

                // Store state back to database

                currentState.MessageState = this.StateSerializer.Serialize(messageData);
                currentState.LastUpdate = message.CreatedUtc;
                this.StateDatabase.UpdateEntityState(currentState);

                // Generate message

                outputXml = this.MessageProcessor.CreateMessageFromState(messageDefinition, messageData);


                // Validate Message if required

                IMessageValidator validator = messageDefinition as IMessageValidator;
                if (validator != null)
                {
                    var validationResults = validator.ValidateMessage(outputXml);

                    if (validationResults.IsValid == false)
                    {
                        this.Logger.Error(
                            String.Format("Generated message fails message internal message validation - {0}", validationResults.ValidationMessage ?? ""));
                        FusionLogger.InfoMessageNonTransactional(message, Messages.General.FusionLogLevel.Error,
                            String.Format("Constructed message does not conform to internal message schema - will not be republished - {0}",
                            validationResults.ValidationMessage ?? ""));

                        throw new InvalidOperationException(
                            String.Format("Constructed message does not conform to internal message schema - will not be republished - {0}",
                            validationResults.ValidationMessage ?? ""));
                            
                    }
                }


            }
            else
            {
                this.Logger.Info("Configured not to store state");

                outputXml = message.Xml;
            }

            // Blind republish
            Bus.Publish(new Destination
            {
                CreatedUtc = message.CreatedUtc,
                EntityRef = message.EntityRef,
                PrimaryEntityRef = message.PrimaryEntityRef,
                Id = message.Id,
                Originator = message.Originator,
                SchemaVersion = message.SchemaVersion,
                Community = message.Community,
                Xml = outputXml
            }
            );

        }
    }
}
