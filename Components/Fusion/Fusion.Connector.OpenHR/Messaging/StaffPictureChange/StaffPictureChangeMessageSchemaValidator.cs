using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffPictureChange
{
    public class StaffPictureChangeMessageSchemaValidator: SchemaValidatorOutboundFilterHandler<StaffPictureChangeRequest>
    {
        public override bool Handle(StaffPictureChangeRequest message)
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

