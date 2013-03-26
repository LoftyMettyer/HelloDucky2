using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffChange
{
    public class StaffChangeMessageSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffChangeRequest>
    {
        public override bool Handle(StaffChangeRequest message)
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
