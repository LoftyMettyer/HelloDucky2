namespace Fusion.Messages.General
{
    using System;
    using NServiceBus;


    /// <summary>
    /// A basic fusion message, all fusion messages should derive from this
    /// </summary>
    [Serializable]
    public abstract class FusionMessage : IMessage
    {

        /// <summary>
        /// Gets or sets the message identifier - this uniquely identifies this message on the bus, ever
        /// </summary>
        /// <value>
        /// The identifier.
        /// </value>
        public Guid Id { get; set; }

        /// <summary>
        /// Gets or sets the originator of the message
        /// </summary>
        /// <value>
        /// The originator.
        /// </value>
        public string Originator { get; set; }


        /// <summary>
        /// Gets or sets the xml message text. This must confirm to the schema of the given message type
        /// </summary>
        /// <value>
        /// The xml.
        /// </value>
        public string Xml { get; set; }


        /// <summary>
        /// Gets or sets the Date/Time of the created.  This must be in UTC
        /// </summary>
        /// <value>
        /// The created.
        /// </value>
        public DateTime CreatedUtc { get; set; }

        /// <summary>
        /// Gets or sets the version of the schema in use
        /// </summary>
        /// <value>
        /// The version.
        /// </value>
        public int SchemaVersion { get; set; }

        /// <summary>
        /// Gets or sets the bus-reference of the entity being discussed, if appropriate for this message type.
        /// </summary>
        /// <value>
        /// The entity reference.
        /// </value>
        public Guid? EntityRef { get; set; }
    }
}