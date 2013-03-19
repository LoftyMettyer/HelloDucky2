using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffContractChange
{
    public class StaffContractChangeMessageSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffContractChangeRequest>
    {
        public override bool Handle(StaffContractChangeRequest message)
        {
            var valid = CheckValidity(message);
            return valid;
        }
    }
}
