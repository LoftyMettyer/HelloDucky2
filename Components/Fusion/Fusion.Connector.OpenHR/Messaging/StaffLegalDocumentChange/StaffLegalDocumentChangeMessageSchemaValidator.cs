using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffLegalDocumentChange
{
    public class StaffLegalDocumentChangeMessageSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffLegalDocumentChangeRequest>
    {
        public override bool Handle(StaffLegalDocumentChangeRequest message)
        {
            var valid = CheckValidity(message);
            return valid;
        }
    }
}
