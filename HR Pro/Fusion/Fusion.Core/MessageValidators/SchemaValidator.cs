// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SchemaValidator.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the schema validator class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Core.MessageValidators
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using System.Xml.Schema;


    /// <summary>
    /// Schema validator. 
    /// </summary>
    public class SchemaValidator
    {
        /// <summary>
        /// Initializes a new instance of the SchemaValidator class.
        /// </summary>
        /// <param name="ns">      The namespace for the schemas. </param>
        /// <param name="baseUri"> URI of the base. </param>
        public SchemaValidator(string ns, string baseUri)
        {
            this.XmlResolver = null;

            this.SchemaNamespace = ns;
            this.BaseSchemaUri = baseUri;
        }

        /// <summary>
        /// Initializes a new instance of the SchemaValidator class.
        /// </summary>
        /// <param name="resolver"> A non-default XmlResolver to use to find schemas. </param>
        /// <param name="ns">       The namespace for the schemas. </param>
        /// <param name="baseUri">  URI of the base. </param>
        public SchemaValidator(XmlResolver resolver, string ns, string baseUri)
        {
            this.XmlResolver = resolver;

            this.SchemaNamespace = ns;
            this.BaseSchemaUri = baseUri;
        }

        /// <summary>
        /// Gets or sets the schema namespace.
        /// </summary>
        /// <value>
        /// The schema namespace.
        /// </value>
        public string SchemaNamespace
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets URI of the base schema.
        /// </summary>
        /// <value>
        /// The base schema uri.
        /// </value>
        public string BaseSchemaUri
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets the xml resolver.
        /// </summary>
        /// <value>
        /// The xml resolver.
        /// </value>
        public XmlResolver XmlResolver
        {
            get;
            private set;
        }

        /// <summary>
        /// Validates the message text xml against the schema. The schema is a uri located relative to the base URL passed in the constructor
        /// </summary>
        /// <param name="messageText"> The message text. </param>
        /// <param name="schemaUrl"> URL of the schema to validate against. </param>
        /// <returns>
        /// Schema validation results
        /// </returns>
        public SchemaValidationResults Validate(string messageText, string schemaUrl)
        {
            XmlReaderSettings readerSettings = new XmlReaderSettings();
            if (this.XmlResolver != null)
            {
                readerSettings.XmlResolver = this.XmlResolver;
                readerSettings.Schemas.XmlResolver = this.XmlResolver;
            }

            readerSettings.Schemas.Add(this.SchemaNamespace, new Uri(new Uri(BaseSchemaUri), schemaUrl).ToString());

            readerSettings.ValidationType = ValidationType.Schema;
            readerSettings.ValidationFlags = XmlSchemaValidationFlags.ReportValidationWarnings;
            readerSettings.ValidationEventHandler += new ValidationEventHandler(ValidationEventHandler);

            validationErrors = new List<string>();

            using (XmlReader xmlReader = XmlReader.Create(new StringReader(messageText), readerSettings))
            {
                try
                {
                    while (xmlReader.Read()) { }
                }
                catch (XmlException)
                {
                    // deal with error here
                }
            }
          
            return new SchemaValidationResults
            {
                ValidationErrors = validationErrors.ToArray()
            };
        }

        /// <summary> The validation errors </summary>
        private List<string> validationErrors = new List<string>();

        /// <summary>
        /// Validation event handler. Collects the schema validation failures
        /// </summary>
        /// <param name="sender"> Source of the event. </param>
        /// <param name="e">      Validation event information. </param>
        private void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            this.validationErrors.Add(e.Exception.Message);
        }

    }
}
