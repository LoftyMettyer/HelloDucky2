using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Messages.SocialCare.Schemas.Patterns
{
    /// <summary>
    /// This interface is implemented only by the DTOs that when serialized can be put directly into the Xml property of the NServiceBus FussionMesssage.
    /// </summary>
    public interface IFusionXmlMessage
        : IFusionXmlDto
    {
        /// <summary>
        /// Schema version.
        /// </summary>
        int version
        {
            get;
            set;
        }
    }
}
