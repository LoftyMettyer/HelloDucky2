using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Connector.OpenHR.MessageHandlers;
using Fusion.Messages.SocialCare;
using NServiceBus;

namespace Fusion.Connector.OpenHR.MessageHandlers
{
	public class StaffTimeSheetPerContractSubmissionHandler : BaseMessageHandler, IHandleMessages<StaffTimeSheetPerContractSubmissionMessage>
	{
		public void Handle(StaffTimeSheetPerContractSubmissionMessage message)
		{




			throw new NotImplementedException();
		}
	}
}
