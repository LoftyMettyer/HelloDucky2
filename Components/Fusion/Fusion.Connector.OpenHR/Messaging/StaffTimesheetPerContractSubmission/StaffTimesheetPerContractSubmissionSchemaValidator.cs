using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffTimesheetPerContractSubmission
{
    public class StaffTimesheetPerContractSubmissionSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffTimeSheetPerContractSubmissionMessage>
    {
        public override bool Handle(StaffTimeSheetPerContractSubmissionMessage message)
        {
            var valid = CheckValidity(message);
            return valid;
        }
    }
}


