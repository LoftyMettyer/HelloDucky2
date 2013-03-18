using NServiceBus;
using System;

namespace Fusion.Messages.General
{
    [Serializable]
    public class LogMessage : ICommand
    {
        /// <summary>
        /// Gets or sets the message identifier - this uniquely identifies this message on the bus, ever
        /// </summary>
        /// <value>
        /// The identifier.
        /// </value>
        public Guid Id { get; set; }

        public string Source { get; set; }

        public Guid? MessageId { get; set; }

        public Guid? EntityRef { get; set; }

        public string MessageName { get; set; }
        /// <summary>
        /// Gets or sets the primary entity reference - where this is a message that is linked to a master entity
        /// </summary>
        /// <value>
        /// The primary entity reference.
        /// </value>
        public Guid? PrimaryEntityRef { get; set; }
        
        public DateTime TimeUtc { get; set; }
                
        public FusionLogLevel LogLevel { get; set; }

        public string Message { get; set; }

        public string MessageDescription { get; set; }

        public string Community
        {
            get;
            set;
        }
    }

}
