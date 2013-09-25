﻿using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Dapper;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Connector.OpenHR.MessageComponents.Data;
using Fusion.Core;
using Fusion.Messages.SocialCare;
using NServiceBus;

namespace Fusion.Connector.OpenHR.MessageHandlers
{
	public class StaffTimeSheetPerContractSubmissionHandler : BaseMessageHandler, IHandleMessages<StaffTimeSheetPerContractSubmissionMessage>
	{
		public void Handle(StaffTimeSheetPerContractSubmissionMessage message)
		{

			var shouldProcess = StartHandlingMessage(message);
			var isValid = true;

			if (!shouldProcess) return;

			StaffTimesheetPerContractSubmission timesheet;
			using (var sr = new StringReader(message.Xml))
			{
				using (var xr = new XmlTextReader(sr))
				{
					var serializer = new XmlSerializer(typeof(StaffTimesheetPerContractSubmission));
					timesheet = (StaffTimesheetPerContractSubmission)serializer.Deserialize(xr);
				}
			}

			var timesheetRef = new Guid(message.EntityRef.ToString());

			var localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Timesheet, timesheetRef);
			var staffId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, new Guid(message.PrimaryEntityRef.ToString()));

			var isNew = (localId == null);


			SqlParameter idParameter;
			using (var c = new SqlConnection(ConnectionString))
			{
				c.Open();

				var cmd = new SqlCommand("fusion.pMessageUpdate_StaffTimesheetSubmission", c)
				{
					CommandType = CommandType.StoredProcedure
				};

				idParameter = cmd.Parameters.Add(new SqlParameter("@ID", localId ?? (object)DBNull.Value));
				idParameter.SqlDbType = SqlDbType.Int;
				idParameter.Direction = ParameterDirection.InputOutput;

				cmd.Parameters.Add(new SqlParameter("@staffId", staffId ?? (object)DBNull.Value));
				cmd.Parameters.Add(new SqlParameter("@recordIsInactive", timesheet.data.recordStatus));
				cmd.Parameters.Add(new SqlParameter("@timesheetdate", timesheet.data.staffTimesheetPerContract.timesheetDate ?? (object)DBNull.Value));
				cmd.Parameters.Add(new SqlParameter("@plannedHours", timesheet.data.staffTimesheetPerContract.plannedHours ?? (object)DBNull.Value));
				cmd.Parameters.Add(new SqlParameter("@workedHours", timesheet.data.staffTimesheetPerContract.workedHours ?? (object)DBNull.Value));
				cmd.Parameters.Add(new SqlParameter("@toilHoursAccrued", timesheet.data.staffTimesheetPerContract.toilHoursAccrued ?? (object)DBNull.Value));
				cmd.Parameters.Add(new SqlParameter("@holidayHoursTaken", timesheet.data.staffTimesheetPerContract.holidayHoursTaken ?? (object)DBNull.Value));
				cmd.Parameters.Add(new SqlParameter("@toilHoursTaken", timesheet.data.staffTimesheetPerContract.toilHoursTaken ?? (object)DBNull.Value));

				try
				{
					c.Execute("fusion.pSetFusionContext", new { MessageType = message.GetMessageName() }, commandType: CommandType.StoredProcedure);
					cmd.ExecuteNonQuery();
				}
				catch (Exception e)
				{
					Logger.ErrorFormat("Inbound message {0}/{1} - {2} failed database save with error", message.GetMessageName(), message.EntityRef, e.Message);
					isValid = false;
				}

			}

			if (isNew & isValid)
			{
				BusRefTranslator.SetBusRef(EntityTranslationNames.Timesheet, idParameter.Value.ToString(), timesheetRef);
			}  

		}
	}
}
