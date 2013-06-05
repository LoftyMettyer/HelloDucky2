using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NServiceBus;

namespace Fusion.Messages.General
{
    [Serializable]
    public class LogTranslationMessage : ICommand
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

        public DateTime TimeUtc { get; set; }

        public string EntityName { get; set; }

        public Guid BusRef { get; set; }

        public string LocalId { get; set; }

    }
}
