
namespace Fusion.Republisher.Core
{
    using Fusion.Core;
    using Fusion.Core.MessageValidators;

    public class FusionSchemaValidator
    {
        public bool CheckValidity(string xml, 
            string ns,
            string baseUri, string schemaName)
        {
            SchemaValidator sv = new SchemaValidator(
                new EmbeddedXmlResourceResolver(),
                ns, baseUri);
//                "http://advancedcomputersoftware.com/xml/fusion",
//                "res://ExampleConnector/ExampleConnector/Schemas/");

            var validation = sv.Validate(xml, schemaName);

            validationMessage = validation.ValidationErrorString;

            return !validation.HasErrors;
        }

        string validationMessage = null;
        public string ValidationMessage
        {
            get
            {
                return validationMessage;
            }
        }
    }
}
