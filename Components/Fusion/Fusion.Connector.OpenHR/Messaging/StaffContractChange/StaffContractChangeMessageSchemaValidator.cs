using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffContractChange
{
    public class StaffContractChangeMessageSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffContractChangeRequest>
    {
        public override bool Handle(StaffContractChangeRequest message)
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
