using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Dapper;
using Fusion.Connector.OpenHR.Database;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Core;
using Fusion.Messages.SocialCare;
using NServiceBus;


namespace Fusion.Connector.OpenHR.MessageHandlers
{
    public class StaffContractChangeMessageHandler : BaseMessageHandler, IHandleMessages<StaffContractChangeMessage>
    {
        public void Handle(StaffContractChangeMessage message)
        {
            
            bool shouldProcess = StartHandlingMessage(message);
            bool isValid = true;

            if (!shouldProcess) return;

            StaffContractChange contract;
            using (var sr = new StringReader(message.Xml))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    var serializer = new XmlSerializer(typeof(StaffContractChange));
                    contract = (StaffContractChange)serializer.Deserialize(xr);

                }
            }


            var contractRef = new Guid(message.EntityRef.ToString());
	        var parentRef = message.PrimaryEntityRef;

            var localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Contract, contractRef);
            var staffId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, new Guid(message.PrimaryEntityRef.ToString()));

            var isNew = (localId == null);

						if (staffId == null)
						{
							this.Bus().HandleCurrentMessageLater();
							return;
						}


            SqlParameter idParameter;
            using (var c = new SqlConnection(ConnectionString))
            {
                c.Open();

								var original = DatabaseAccess.readContract(Convert.ToInt32(localId));
								var update = contract.data.staffContract;

								// Merge with original if nodes omitted
								if (original != null)
								{
									update.costCenter = !update.costCenterSpecified ? original.costCenter : update.costCenter;
									update.effectiveTo = !update.effectiveToSpecified ? original.effectiveTo : update.effectiveTo;
								}

                var cmd = new SqlCommand("fusion.pMessageUpdate_StaffContractChange", c)
                    {
                        CommandType = CommandType.StoredProcedure
                    };

                idParameter = cmd.Parameters.Add(new SqlParameter("@ID", localId ?? (object)DBNull.Value));
                idParameter.SqlDbType = SqlDbType.Int;
                idParameter.Direction = ParameterDirection.InputOutput;

                cmd.Parameters.Add(new SqlParameter("@staffId", staffId ));
                cmd.Parameters.Add(new SqlParameter("@recordIsInactive", contract.data.recordStatus));
                cmd.Parameters.Add(new SqlParameter("@contractName", contract.data.staffContract.contractName ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@department", contract.data.staffContract.department ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@primarySite", contract.data.staffContract.primarySite ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@contractedHoursPerWeek", contract.data.staffContract.contractedHoursPerWeek ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@maximumHoursPerWeek", contract.data.staffContract.maximumHoursPerWeek ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@effectiveFrom", contract.data.staffContract.effectiveFrom ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@effectiveTo", contract.data.staffContract.effectiveTo ?? (object)DBNull.Value));
								cmd.Parameters.Add(new SqlParameter("@costCenter", contract.data.staffContract.costCenter ?? (object)DBNull.Value));

                try
                {
                    c.Execute("fusion.pSetFusionContext", new { MessageType = message.GetMessageName() }, commandType: CommandType.StoredProcedure);
                    cmd.ExecuteNonQuery();

										//// Store the message in a format as if we'd generated it.
										//var newData = DatabaseAccess.readContract(Convert.ToInt32(localId));
										//var ChangeMessage = new StaffContractChange(contractRef, parentRef, newData);
										MessageTracking.SetLastGeneratedXml(message.GetMessageName(), message.EntityRef.Value, contract.ToXml());

                }
                catch (Exception e)
                {
                    Logger.ErrorFormat("Inbound message {0}/{1} - {2} failed database save with error", message.GetMessageName(), message.EntityRef, e.Message);
                    isValid = false;
                }

            }

            if (isNew & isValid)
            {
                BusRefTranslator.SetBusRef(EntityTranslationNames.Contract, idParameter.Value.ToString(), contractRef);
            }  

        }
    }
}
