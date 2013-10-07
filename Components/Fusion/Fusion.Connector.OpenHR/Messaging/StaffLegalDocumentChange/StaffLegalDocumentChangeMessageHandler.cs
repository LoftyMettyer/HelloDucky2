using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Dapper;
using Fusion.Connector.OpenHR.Database;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using Fusion.Core;
using Fusion.Messages.SocialCare;
using NServiceBus;

namespace Fusion.Connector.OpenHR.MessageHandlers
{
    public class StaffLegalDocumentChangeMessageHandler : BaseMessageHandler, IHandleMessages<StaffLegalDocumentChangeMessage>
    {

        public void Handle(StaffLegalDocumentChangeMessage message)
        {

            bool shouldProcess = StartHandlingMessage(message);
            bool isValid = true;

            if (!shouldProcess) return;

            StaffLegalDocumentChange document;
            using (var sr = new StringReader(message.Xml))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    var serializer = new XmlSerializer(typeof (StaffLegalDocumentChange));
                    document = (StaffLegalDocumentChange)serializer.Deserialize(xr);

                }
            }


            var docRef = new Guid(message.EntityRef.ToString());
						var parentRef = message.PrimaryEntityRef;

            var localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Document, docRef);
	        var staffId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, new Guid(message.PrimaryEntityRef.ToString()));

					Logger.InfoFormat("Inbound StaffLegalDocumentChangeMessage message staffref- {0}, legalRef {1}", message.PrimaryEntityRef.ToString(), message.EntityRef.ToString());

            var isNew = (localId == null && document.data.recordStatus == RecordStatusStandard.Active);

						if (staffId == null)
						{
							var dummyStaff = new Staff
							{
								surname = "** Unknown Fusion **",
								forenames = "** From LegalDocument **",
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
							Logger.InfoFormat("Inbound Created dummy staff record for staffef- {0}, legalRef {1}", message.PrimaryEntityRef.ToString(), message.EntityRef.ToString());

							this.Bus().HandleCurrentMessageLater();
							return;
						}


            SqlParameter idParameter;
            using (var c = new SqlConnection(ConnectionString))
            {
                c.Open();


								var original = DatabaseAccess.readDocument(Convert.ToInt32(localId));
								var update = document.data.staffLegalDocument;

							// Merge with original if nodes omitted
	            if (original != null)
	            {
		            update.acceptedBy = !update.acceptedBySpecified ? original.acceptedBy : update.acceptedBy;
		            update.acceptedDate = !update.acceptedDateSpecified ? original.acceptedDate : update.acceptedDate;
		            update.requestedBy = !update.requestedBySpecified ? original.requestedBy : update.requestedBy;
		            update.requestedDate = !update.requestedDateSpecified ? original.requestedDate : update.requestedDate;
		            update.validFrom = !update.validFromSpecified ? original.validFrom : update.validFrom;
		            update.validTo = !update.validToSpecified ? original.validTo : update.validTo;
	            }

	            var cmd = new SqlCommand("fusion.pMessageUpdate_StaffLegalDocumentChange", c)
                    {
                        CommandType = CommandType.StoredProcedure
                    };

                idParameter = cmd.Parameters.Add(new SqlParameter("@ID", localId ?? (object)DBNull.Value));
                idParameter.SqlDbType = SqlDbType.Int;
                idParameter.Direction = ParameterDirection.InputOutput;

                cmd.Parameters.Add(new SqlParameter("@staffId", staffId));
                cmd.Parameters.Add(new SqlParameter("@recordIsInactive", document.data.recordStatus));
                cmd.Parameters.Add(new SqlParameter("@typeName", document.data.staffLegalDocument.typeName.ToString()));
                cmd.Parameters.Add(new SqlParameter("@validFrom", document.data.staffLegalDocument.validFrom ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@validTo", document.data.staffLegalDocument.validTo ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@documentReference", document.data.staffLegalDocument.documentReference ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@requestedBy", document.data.staffLegalDocument.requestedBy ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@requestedDate", document.data.staffLegalDocument.requestedDate ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@acceptedBy", document.data.staffLegalDocument.acceptedBy ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@acceptedDate", document.data.staffLegalDocument.acceptedDate ?? (object)DBNull.Value));

                try
                {
                    c.Execute("fusion.pSetFusionContext", new { MessageType = message.GetMessageName() }, commandType: CommandType.StoredProcedure);
                    cmd.ExecuteNonQuery();

										//// Store the message in a format as if we'd generated it.
										//var newData = DatabaseAccess.readDocument(Convert.ToInt32(localId));
										//var ChangeMessage = new StaffLegalDocumentChange(docRef, parentRef, newData);
										MessageTracking.SetLastGeneratedXml(message.GetMessageName(), message.EntityRef.Value, document.ToXml());

                }
                catch (Exception e)
                {
                    Logger.ErrorFormat("Inbound message {0}/{1} - {2} failed database save with error", message.GetMessageName(), message.EntityRef, e.Message);
                    isValid = false;
                }

            }

            if (isNew & isValid)
            {
                BusRefTranslator.SetBusRef(EntityTranslationNames.Document, idParameter.Value.ToString(), docRef);
            }
        }
    }
}
