using Nexus.Common.Interfaces;
using OpenHRNexus.Common.Enums;
using Nexus.Common.Enums;

namespace Nexus.Common.Classes {
	public class BusinessProcessStepEmail : IBusinessProcessStep {
		public int Id { get; set; }

		public BusinessProcessStepType Type {
			get {
				return BusinessProcessStepType.Email;
			}
		}

		public BusinessProcessStepStatus Validate() {
			return BusinessProcessStepStatus.Success;
		}

		public string To => "roberto.caballero@advancedcomputersoftware.com";

		public string Message =>
			"<!DOCTYPE html>" +
			"<html lang='en'>" +
			"    <head>" +
			"        <meta charset='utf-8' />" +
			"    </head>" +
			"    <body>" +
			"        <p>" +
			"            <span style='color: #0094ff'>{0}</span> has requested a <span style='color:#0094ff'>{1}</span> holiday absence from <span style='color:#0094ff'>{2}</span> to <span style='color:#0094ff'>{3}.</span>" +
			"        </p>" +
			"        <p>" +
			"            Reason for absence: <span style='color: #0094ff'>{4}</span>" +
			"        </p>" +
			"        <p>" +
			"            Employee notes: <span style='color: #0094ff'>{5}</span>" +
			"        </p>" +
			"        <p>" +
			"            You can quickly approve or decline this absence request using the buttons below." +
			"        </p>" +
			"        <span style='background: green; padding: 5px'><a style='text-decoration: none; color: white' href='{6}'>Approve</a></span>" +
			"        <span style='background: red; padding: 5px'><a style='text-decoration: none; color: white' href='{7}'>Decline</a></span>" +
			"        <span style='background: lightblue; padding: 5px'><a style='text-decoration: none; color: white' href='{7}'>View the request</a></span>" +
			"        <span style='background: blue; padding: 5px'><a style='text-decoration: none; color: white' href='{9}'>View team calendar</a></span>" +
			"    </body>" +
			"</html>";

		public string Subject => "Nexus subject";
	}
}
