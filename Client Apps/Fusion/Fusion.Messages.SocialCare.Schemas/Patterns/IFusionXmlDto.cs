using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Messages.SocialCare.Schemas.Patterns
{
    /// <summary>
    /// This interface is implemented by all the DTOs generated from Fusion XSD.
    /// </summary>
    public interface IFusionXmlDto
    {
        string Serialize();
    }
}
