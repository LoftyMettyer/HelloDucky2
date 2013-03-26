using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffContactChange
{
    public class StaffContactChangeMessageSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffContactChangeRequest>
    {
        public override bool Handle(StaffContactChangeRequest message)
        {
            var valid = false;

            if (!checkAlreadySent(message))
            {
                valid = CheckValidity(message);
            }

            return valid;
        }

    }
}
