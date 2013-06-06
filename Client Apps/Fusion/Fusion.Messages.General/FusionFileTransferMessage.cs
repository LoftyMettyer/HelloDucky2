namespace Fusion.Messages.General
{
    using System;
    using NServiceBus;


    /// <summary>
    /// A basic fusion message, all fusion messages should derive from this
    /// </summary>
    [Serializable]
    public abstract class FusionFileTransferMessage : ICommand
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
        /// Gets or sets the filename of the file.
        /// </summary>
        /// <value>
        /// The filename.
        /// </value>
        public string Filename { get; set; }

        /// <summary>
        /// Gets or sets the file text. This must confirm to the schema of the given message type
        /// </summary>
        /// <value>
        /// The xml.
        /// </value>
        public byte[] Xml { get; set; }


        /// <summary>
        /// Gets or sets the Date/Time of the message creation.  This must be in UTC
        /// </summary>
        /// <value>
        /// The created.
        /// </value>
        public DateTime Created { get; set; }

        /// <summary>
        /// Gets or sets the Date/Time of the file modification.  This must be in UTC
        /// </summary>
        /// <value>
        /// The created.
        /// </value>
        public DateTime LastModifiedDate { get; set; }

    }
}