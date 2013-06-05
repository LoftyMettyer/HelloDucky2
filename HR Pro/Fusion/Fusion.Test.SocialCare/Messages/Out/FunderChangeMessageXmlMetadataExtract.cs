// --------------------------------------------------------------------------------------------------------------------
// <copyright file="StaffChangeRequestXmlMetadataExtract.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the staff change request xml metadata extract class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Test.SocialCare.Messages
{
    using System;
    using System.IO;
    using System.Xml.Linq;
    using Fusion.Messages.SocialCare;
    using Fusion.Core.Test;

    /// <summary>
    /// Extract entity reference number and version number from message
    /// </summary>
    public class FunderChangeMessageXmlMetadataExtract : FusionXmlMetadataExtract<FunderChangeMessage>
    {
        /// <summary>
        /// Gets an xml metadata.
        /// </summary>
        /// <param name="message"> The message. </param>
        /// <returns>
        /// The entity reference.
        /// </returns>
        public override FusionXmlMetadata GetMetadataFromXml(FunderChangeMessage message)
        {
            XNamespace ns = "http://advancedcomputersoftware.com/xml/fusion/socialCare";
            XDocument fusionDocument = XDocument.Load(new StringReader(message.Xml));

            var rootNode = fusionDocument.Element(ns + "funderChange");

            string entityRef = (string)rootNode.Attribute("funderRef");
            string version = (string)rootNode.Attribute("version");

            return new FusionXmlMetadata
            {
                EntityRef = new Guid(entityRef),
                Version = Convert.ToInt32(version)
            };
        }
    }
}
