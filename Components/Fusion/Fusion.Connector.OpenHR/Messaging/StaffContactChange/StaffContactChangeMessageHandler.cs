using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Dapper;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Core;
using Fusion.Messages.SocialCare;
using NServiceBus;
using Address = Fusion.Connector.OpenHR.MessageComponents.Component.Address;

namespace Fusion.Connector.OpenHR.MessageHandlers
{
    public class StaffContactChangeMessageHandler : BaseMessageHandler, IHandleMessages<StaffContactChangeMessage>
    {
        public void Handle(StaffContactChangeMessage message)
        {

            var shouldProcess = StartHandlingMessage(message);
            var isValid = true;

            if (!shouldProcess) return;

            StaffContactChange contact;
            using (var sr = new StringReader(message.Xml))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    var serializer = new XmlSerializer(typeof(StaffContactChange));
                    contact = (StaffContactChange)serializer.Deserialize(xr);

                }
            }


            var contactRef = new Guid(message.EntityRef.ToString());
	        var parentRef = message.PrimaryEntityRef;

            var localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Contact, contactRef);
            var staffId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, new Guid(message.PrimaryEntityRef.ToString()));

						Logger.InfoFormat("Inbound StaffContactChangeMessageHandler message staffref- {0}, contactRef {1}", message.PrimaryEntityRef.ToString(), message.EntityRef.ToString());

            var isNew = (localId == null);

						if (staffId == null)
						{
							var dummyStaff = new Staff
							{
								surname = "** Unknown Fusion **",
								forenames = "** From StaffContact **",
								email = message.PrimaryEntityRef.ToString()
							};
							var dummyStaffChange = new StaffChange(new Guid(message.PrimaryEntityRef.ToString()), dummyStaff);
							var dummyStaffMessage = new StaffChangeMessage
							{
								Community = message.Community,
								CreatedUtc = DateTime.Now,
								EntityRef = message.PrimaryEntityRef,
								Originator = message.Originator,
								SchemaVersion = message.SchemaVersion,
								Xml = dummyStaffChange.ToXml()
							};

							var handler = new StaffChangeMessageHandler
							{
								BusRefTranslator = BusRefTranslator,
								MessageTracking = MessageTracking
							};
							handler.SaveToDB(dummyStaffChange, dummyStaffMessage);
							Logger.InfoFormat("Inbound Created dummy staff record for staffref- {0}, contactRef {1}", message.PrimaryEntityRef.ToString(), message.EntityRef.ToString());

							this.Bus().HandleCurrentMessageLater();
							return;
						}

            SqlParameter idParameter;
            using (var c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var cmd = new SqlCommand("fusion.pMessageUpdate_StaffContactChange", c)
                {
                    CommandType = CommandType.StoredProcedure
                };

                idParameter = cmd.Parameters.Add(new SqlParameter("@ID", localId ?? (object)DBNull.Value));
                idParameter.SqlDbType = SqlDbType.Int;
                idParameter.Direction = ParameterDirection.InputOutput;

                cmd.Parameters.Add(new SqlParameter("@staffId", staffId ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@recordIsInactive", contact.data.recordStatus));

                cmd.Parameters.Add(new SqlParameter("@title", contact.data.staffContact.title ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@forenames", contact.data.staffContact.forenames ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@surname", contact.data.staffContact.surname ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@contactType", contact.data.staffContact.contactType ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@relationshipType", contact.data.staffContact.relationshipType ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@workMobile", contact.data.staffContact.workMobile ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@personalMobile",  contact.data.staffContact.personalMobile ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@workPhoneNumber",  contact.data.staffContact.workPhoneNumber ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@homePhoneNumber", contact.data.staffContact.homePhoneNumber ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@email", contact.data.staffContact.email ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@notes", contact.data.staffContact.notes ?? (object)DBNull.Value));

                if (contact.data.staffContact.homeAddress == null)
                {
                    contact.data.staffContact.homeAddress = new Address();
                }

                cmd.Parameters.Add(new SqlParameter("@addressline1", contact.data.staffContact.homeAddress.addressLine1 ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@addressline2", contact.data.staffContact.homeAddress.addressLine2 ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@addressline3", contact.data.staffContact.homeAddress.addressLine3 ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@addressline4", contact.data.staffContact.homeAddress.addressLine4 ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@addressline5", contact.data.staffContact.homeAddress.addressLine5 ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@postcode", contact.data.staffContact.homeAddress.postCode ?? (object)DBNull.Value));

                try
                {
                    c.Execute("fusion.pSetFusionContext", new { MessageType = message.GetMessageName() }, commandType: CommandType.StoredProcedure);
                    cmd.ExecuteNonQuery();

										//// Store the message in a format as if we'd generated it.
										//var newData = DatabaseAccess.readContact(Convert.ToInt32(localId));
										//var ChangeMessage = new StaffContactChange(contactRef, parentRef, newData);
										MessageTracking.SetLastGeneratedXml(message.GetMessageName(), message.EntityRef.Value, contact.ToXml());

										if (isNew & isValid)
										{
											BusRefTranslator.SetBusRef(EntityTranslationNames.Contact, idParameter.Value.ToString(), contactRef);
										}  

                }
                catch (Exception e)
                {
                    Logger.ErrorFormat("Inbound message {0}/{1} - {2} failed database save with error", message.GetMessageName(), message.EntityRef, e.Message);
										this.Bus().HandleCurrentMessageLater();
                    isValid = false;
                }

            }
        
        }
    }
}
