using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Dapper;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageHandlers;
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

            var localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Contact, contactRef);
            var staffId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, new Guid(message.PrimaryEntityRef.ToString()));

            var isNew = (localId == null);


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
                cmd.Parameters.Add(new SqlParameter("@description", contact.data.staffContact.description ?? (object)DBNull.Value));
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
                }
                catch (Exception e)
                {
                    Logger.ErrorFormat("Inbound message {0}/{1} - {2} failed database save with error", message.GetMessageName(), message.EntityRef, e.Message);
                    isValid = false;
                }

            }

            if (isNew & isValid)
            {
                BusRefTranslator.SetBusRef(EntityTranslationNames.Contact, idParameter.Value.ToString(), contactRef);
            }  
        
        
        
        
        }
    }
}
