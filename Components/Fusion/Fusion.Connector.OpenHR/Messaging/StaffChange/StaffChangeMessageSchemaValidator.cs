using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffChange
{
    public class StaffChangeMessageSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffChangeRequest>
    {
        public override bool Handle(StaffChangeRequest message)
        {
            var valid = CheckValidity(message);
            return valid;
        }
    }
}
