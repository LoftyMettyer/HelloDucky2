using System;
using System.Collections.Specialized;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using Nexus.Common.Classes;
using Nexus.Common.Enums;
using Nexus.EnterpriseService.ExceptionHandling;
using Nexus.Service.Globals;
using PostmarkDotNet;
using PostmarkDotNet.Legacy;

namespace Nexus.Service {
	public class EmailService {
		public BusinessProcessStepResponse Send(string to, string subject, string body) {
			var message = new PostmarkMessage {
				From = "roberto.caballero@advancedcomputersoftware.com",
				To = to,
				Subject = subject,
				HtmlBody = body,
				TextBody = body,
				ReplyTo = "nexus-reply@advancedcomputersoftware.com"
			};

			var client = new PostmarkClient("4984ad83-2881-46ee-998c-97b0523822df");

			var response = client.SendMessage(message);

			if (response.Status != PostmarkStatus.Success) {
				return new BusinessProcessStepResponse() {
					Status = BusinessProcessStepStatus.EmailFailedToSend,
					Message = "Email failed to send",
					FollowOnUrl = String.Empty
				};
			}

			return new BusinessProcessStepResponse() {
				Status = BusinessProcessStepStatus.EmailSuccessfullySent,
				Message = "Email successfully sent",
				FollowOnUrl = String.Empty
			};
		}
	}
}