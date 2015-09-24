using System;
using Nexus.Common.Classes;
using Nexus.Common.Enums;
using System.Net.Mail;
using PostmarkDotNet;
using PostmarkDotNet.Legacy;

namespace Nexus.Service.Services {
	public class EmailService {
		public ProcessStepResponse Send(MailMessage message) {

			//var sendMessage = new PostmarkMessage
			//{
			//    From = message.From.ToString(),
			//    To = message.To.ToString(),
			//    Subject = message.Subject,
			//    HtmlBody = message.Body,
			//    TextBody = message.Body,
			//    ReplyTo = message.ReplyToList.ToString()
			//};

			//var client = new PostmarkClient("4984ad83-2881-46ee-998c-97b0523822df");

			//var response = client.SendMessage(sendMessage);

			//if (response.Status != PostmarkStatus.Success)
			//{
			//    return new ProcessStepResponse()
			//    {
			//        Status = ProcessStepStatus.EmailFailedToSend,
			//        Message = "Email failed to send",
			//        FollowOnUrl = String.Empty
			//    };
			//}

			//Uncoment the lines above to restore email - sending functionality//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

			SmtpClient smtpClient = new SmtpClient("localhost", 25)
			{
				UseDefaultCredentials = true,
				DeliveryMethod = SmtpDeliveryMethod.Network
			};

			MailMessage mail = new MailMessage();

			mail.From = new MailAddress("nexus@nexus.com", "Nexus");

			foreach (var email in message.To) {
				mail.To.Add(email);
			}

			mail.Subject = message.Subject;
			mail.Body = message.Body;

			smtpClient.Send(mail);

			return new ProcessStepResponse() {
				Status = ProcessStepStatus.EmailSuccessfullySent,
				Message = "Email successfully sent",
				FollowOnUrl = String.Empty
			};


			return new ProcessStepResponse() {
				Status = ProcessStepStatus.EmailSuccessfullySent,
				Message = "Email successfully sent",
				FollowOnUrl = String.Empty
			};
		}

	}
}