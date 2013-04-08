using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffLegalDocumentChange
{
    public class StaffLegalDocumentChangeMessageSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffLegalDocumentChangeRequest>
    {
        public override bool Handle(StaffLegalDocumentChangeRequest message)
        {
            var valid = false;

            if (checkStaffHasBeenSent(message))
            {
                if (!checkAlreadySent(message))
                {
                    valid = CheckValidity(message);
                }
            }

            return valid;
        }
    }
}
