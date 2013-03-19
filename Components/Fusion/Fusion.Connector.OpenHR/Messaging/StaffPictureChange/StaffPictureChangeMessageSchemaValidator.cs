using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffPictureChange
{
    public class StaffPictureChangeMessageSchemaValidator: SchemaValidatorOutboundFilterHandler<StaffPictureChangeRequest>
    {
        public override bool Handle(StaffPictureChangeRequest message)
        {
            var valid = CheckValidity(message);
            return valid;
        }
    }
}

