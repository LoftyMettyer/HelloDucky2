// --------------------------------------------------------------------------------------------------------------------
// <copyright file="GenericBlindMessageRepublisher.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the generic blind message republisher class
// </summary>
// --------------------------------------------------------------------------------------------------------------------


namespace Fusion.Republisher.Core
{
    using Fusion.Core;
    using Fusion.Messages.General;
    using log4net;
    using NServiceBus;
    using StructureMap.Attributes;
    using Fusion.Core.Logging;

    public class GenericBlindMessageRepublisher<Source, Destination> 
        : IHandleMessages<Source>

        where Source : FusionMessage, ICommand                                                                     
        where Destination : FusionMessage, IEvent, new()
    {
        public IBus Bus
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


        public GenericBlindMessageRepublisher()
        {
            this.Logger = LogManager.GetLogger(this.GetType());
        }

        public void Handle(Source message)
        {
            Logger.Info(string.Format("Fusion Performing Blind republish received {0} from {1} with Id {2}.", message.GetMessageName(), message.Originator, message.Id));
            Logger.Info(string.Format("Message time: {0}.", message.CreatedUtc));

            FusionLogger.LogMessageReceived(message);

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
                Xml = message.Xml
            }
            );
        }

        private readonly ILog Logger;
    }


}
