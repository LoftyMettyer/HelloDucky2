﻿using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Dapper;
using Fusion.Connector.OpenHR.Database;
using Fusion.Connector.OpenHR.MessageComponents;
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

            var localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Document, docRef);
            var staffId = Convert.ToInt32(BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, new Guid(message.PrimaryEntityRef.ToString())));

            var isNew = (localId == null && document.data.recordStatus == RecordStatusStandard.Active);

						// Push to saga?
						if (staffId == 0)
						{
							Logger.WarnFormat("Inbound message {0}/{1} pushed to saga. No parent found for {2}", message.GetMessageName(), message.EntityRef, message.PrimaryEntityRef);
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
