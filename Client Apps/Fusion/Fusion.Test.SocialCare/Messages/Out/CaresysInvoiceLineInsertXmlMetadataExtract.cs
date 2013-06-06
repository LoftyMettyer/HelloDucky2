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
    public class CaresysInvoiceLineInsertMessageMetadataExtract : FusionXmlMetadataExtract<CaresysInvoiceLineInsertMessage>
    {
        /// <summary>
        /// Gets an xml metadata.
        /// </summary>
        /// <param name="message"> The message. </param>
        /// <returns>
        /// The entity reference.
        /// </returns>
        public override FusionXmlMetadata GetMetadataFromXml(CaresysInvoiceLineInsertMessage message)
        {
            XNamespace ns = "http://advancedcomputersoftware.com/xml/fusion/socialCare";
            XDocument fusionDocument = XDocument.Load(new StringReader(message.Xml));

            var rootNode = fusionDocument.Element(ns + "careSysInvoiceLineInsert");

            //string entityRef = (string)rootNode.Attribute("fundingRef");
            string version = (string)rootNode.Attribute("version");

            return new FusionXmlMetadata
            {
                EntityRef = null,
                Version = Convert.ToInt32(version)
            };
        }
    }
}
