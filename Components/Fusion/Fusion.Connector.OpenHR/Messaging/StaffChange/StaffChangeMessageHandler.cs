using System.Data.SqlClient;
using System.Linq;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Connector.OpenHR.Database;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Core;
using Fusion.Core.Sql.OutboundBuilder;
using log4net;

using NServiceBus;
using StructureMap.Attributes;
using System.IO;
using System;
using Fusion.Messages.SocialCare;
using Fusion.Connector.OpenHR.MessageComponents;
using System.Xml.Serialization;
using System.Xml;
using Dapper;
using System.Data;
using Address = Fusion.Connector.OpenHR.MessageComponents.Component.Address;

namespace Fusion.Connector.OpenHR.MessageHandlers
{
    public class StaffChangeMessageHandler : BaseMessageHandler, IHandleMessages<StaffChangeMessage>
    {

			[SetterProperty]
			public IOutboundBuilderFactory OutboundBuilderFactory { get; set; }

			[SetterProperty]
			public IFusionConfiguration config { get; set; }

        private readonly string _connectionString;

        public StaffChangeMessageHandler () {
            _connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            Logger = LogManager.GetLogger(typeof(StaffChangeMessageHandler));
        }

				public void Handle(StaffChangeMessage message)
				{

					var shouldProcess = StartHandlingMessage(message);

					if (!shouldProcess) return;

					StaffChange staff;
					using (var sr = new StringReader(message.Xml))
					{
						using (var xr = new XmlTextReader(sr))
						{
							var serializer = new XmlSerializer(typeof(StaffChange));
							staff = (StaffChange)serializer.Deserialize(xr);
						}
					}

					SaveToDB(staff, message);
				}

				public void SaveToDB(StaffChange staff, StaffChangeMessage message)
				{
					var isValid = true;

					SqlParameter idParameter;

					var busRef = new Guid(staff.staffRef);
					var localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, busRef);

					var isNew = (localId == null);

					using (var c = new SqlConnection(_connectionString))
					{
						c.Open();

						var cmd = new SqlCommand("fusion.pMessageUpdate_StaffChange", c)
						{
							CommandType = CommandType.StoredProcedure
						};

						idParameter = cmd.Parameters.Add(new SqlParameter("@ID", localId ?? (object)DBNull.Value));
						idParameter.SqlDbType = SqlDbType.Int;
						idParameter.Direction = ParameterDirection.InputOutput;

						cmd.Parameters.Add(new SqlParameter("@recordIsInactive", staff.data.recordStatus));
						cmd.Parameters.Add(new SqlParameter("@title", staff.data.staff.title ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@forenames", staff.data.staff.forenames ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@surname", staff.data.staff.surname ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@preferredName", staff.data.staff.preferredName ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@payrollNumber", staff.data.staff.payrollNumber ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@DOB", staff.data.staff.dob ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@employeeType", staff.data.staff.employeeType ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@workMobile", staff.data.staff.workMobile ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@personalMobile", staff.data.staff.personalMobile ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@workPhoneNumber", staff.data.staff.workPhoneNumber ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@homePhoneNumber", staff.data.staff.homePhoneNumber ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@email", staff.data.staff.email ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@personalEmail", staff.data.staff.personalEmail ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@gender", staff.data.staff.gender.ToString()));
						cmd.Parameters.Add(new SqlParameter("@startDate", staff.data.staff.startDate ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@leavingDate", staff.data.staff.leavingDate ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@leavingReason", staff.data.staff.leavingReason ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@companyName", staff.data.staff.companyName ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@jobTitle", staff.data.staff.jobTitle ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@managerRef", staff.data.staff.managerRef ?? (object)DBNull.Value));

						if (staff.data.staff.homeAddress == null)
						{
							staff.data.staff.homeAddress = new Address();
						}

						cmd.Parameters.Add(new SqlParameter("@addressLine1", staff.data.staff.homeAddress.addressLine1 ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@addressLine2", staff.data.staff.homeAddress.addressLine2 ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@addressLine3", staff.data.staff.homeAddress.addressLine3 ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@addressLine4", staff.data.staff.homeAddress.addressLine4 ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@addressLine5", staff.data.staff.homeAddress.addressLine5 ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@postcode", staff.data.staff.homeAddress.postCode ?? (object)DBNull.Value));
						cmd.Parameters.Add(new SqlParameter("@nationalInsuranceNumber", staff.data.staff.nationalInsuranceNumber ?? (object)DBNull.Value));

						try
						{
							c.Execute("fusion.pSetFusionContext", new { MessageType = message.GetMessageName() }, commandType: CommandType.StoredProcedure);
							cmd.ExecuteNonQuery();

							//// Store the message in a format as if we'd generated it.
							//var newData = DatabaseAccess.readStaff(Convert.ToInt32(localId));
							//var ChangeMessage = new StaffChange(busRef, newData);
							MessageTracking.SetLastGeneratedXml(message.GetMessageName(), message.EntityRef.Value, staff.ToXml());

							if (isNew & isValid)
							{
								BusRefTranslator.SetBusRef(EntityTranslationNames.Staff, idParameter.Value.ToString(), busRef);
							}			

						}
						catch (Exception e)
						{
							Logger.ErrorFormat("Inbound message {0}/{1} - {2} failed database save with error", message.GetMessageName(), message.EntityRef, e.Message);
							isValid = false;
							throw;
						}

					}

				}

    }

}


