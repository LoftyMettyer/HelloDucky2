    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

namespace ProgressConnector.ProgressInterface
{
    /// <summary>
    /// Content received from progress representing a message to transmit to the bus
    /// </summary>
    public class OpenExchangeMessage : OpenExchangeGeneratedContent
    {

        /// <summary>
        /// Gets or sets the type of the message.
        /// </summary>
        /// <value>
        /// The type of the message.
        /// </value>
        public string MessageType
        {
            get;
            set;
        }
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
        public DateTime Created { get; set; }

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

        public Guid? PrimaryEntityRef { get; set; }

        public string Community { get; set; }
    }
}

