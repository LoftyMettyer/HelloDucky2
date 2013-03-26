using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Dapper;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Connector.OpenHR.MessageHandlers;
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


            var contactRef = new Guid(message.EntityRef.ToString());

            var localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Contract, contactRef);
            var staffId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, new Guid(message.PrimaryEntityRef.ToString()));

            var isNew = (localId == null);


            SqlParameter idParameter;
            using (var c = new SqlConnection(ConnectionString))
            {
                c.Open();

                var cmd = new SqlCommand("fusion.pMessageUpdate_StaffContractChange", c)
                    {
                        CommandType = CommandType.StoredProcedure
                    };

                idParameter = cmd.Parameters.Add(new SqlParameter("@ID", localId ?? (object)DBNull.Value));
                idParameter.SqlDbType = SqlDbType.Int;
                idParameter.Direction = ParameterDirection.InputOutput;

                cmd.Parameters.Add(new SqlParameter("@staffId", staffId ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@contractName", contract.data.staffContract.contractName ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@department", contract.data.staffContract.department ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@primarySite", contract.data.staffContract.primarySite ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@contractedHoursPerWeek", contract.data.staffContract.contractedHoursPerWeek ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@maximumHoursPerWeek", contract.data.staffContract.maximumHoursPerWeek ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@effectiveFrom", contract.data.staffContract.effectiveFrom ?? (object)DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("@effectiveTo", contract.data.staffContract.effectiveTo ?? (object)DBNull.Value));

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
                BusRefTranslator.SetBusRef(EntityTranslationNames.Contract, idParameter.Value.ToString(), contactRef);
            }  

        }
    }
}
