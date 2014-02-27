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
using Fusion.Core;
using Fusion.Messages.SocialCare;
using NServiceBus;

namespace Fusion.Connector.OpenHR.MessageHandlers
{
    public class StaffSkillChangeMessageHandler : BaseMessageHandler, IHandleMessages<StaffSkillChangeMessage>
    {
        public void Handle(StaffSkillChangeMessage message)
        {

            bool shouldProcess = base.StartHandlingMessage(message);
            bool isNew = true;
            bool isValid = true;


            if (shouldProcess == true)
            {
                StaffSkillChange skill;
                using (var sr = new StringReader(message.Xml))
                {
                    using (var xr = new XmlTextReader(sr))
                    {
                        var serializer = new XmlSerializer(typeof(StaffSkillChange));
                        skill = (StaffSkillChange)serializer.Deserialize(xr);

                    }
                }


                var skillRef = new Guid(message.EntityRef.ToString());
                var localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Skill, skillRef);
                var staffId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, new Guid(message.PrimaryEntityRef.ToString()));

								Logger.InfoFormat("Inbound StaffSkillChangeMessage message staffref- {0}, skillref {1}", message.PrimaryEntityRef.ToString(), message.EntityRef.ToString());

                isNew = (localId == null);

								if (staffId == null)
								{
									var dummyStaff = new Staff
									{
										surname = "** Unknown Fusion **",
										forenames = "** From StaffSkill **",
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
									Logger.InfoFormat("Inbound Created dummy staff record for staffref- {0}, skillRef {1}", message.PrimaryEntityRef.ToString(), message.EntityRef.ToString());

									this.Bus().HandleCurrentMessageLater();
									return;
								}


                SqlParameter idParameter;
                using (var c = new SqlConnection(ConnectionString))
                {
                    c.Open();

                    var cmd = new SqlCommand("fusion.pMessageUpdate_StaffSkillChange", c)
                    {
                        CommandType = CommandType.StoredProcedure
                    };

                    idParameter = cmd.Parameters.Add(new SqlParameter("@ID", localId ?? (object)DBNull.Value));
                    idParameter.SqlDbType = SqlDbType.Int;
                    idParameter.Direction = ParameterDirection.InputOutput;

                    cmd.Parameters.Add(new SqlParameter("@staffId", staffId ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@recordIsInactive", skill.data.recordStatus));
                    cmd.Parameters.Add(new SqlParameter("@name", skill.data.staffSkill.name ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@trainingStart", skill.data.staffSkill.trainingStart ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@trainingEnd", skill.data.staffSkill.trainingEnd ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@validFrom", skill.data.staffSkill.validFrom ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@validTo", skill.data.staffSkill.validTo ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@reference", skill.data.staffSkill.reference ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@outcome", skill.data.staffSkill.outcome ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@didNotAttend", skill.data.staffSkill.didNotAttend ?? (object)DBNull.Value));

                    try
                    {
                        c.Execute("fusion.pSetFusionContext", new { MessageType = message.GetMessageName() }, commandType: CommandType.StoredProcedure);
                        cmd.ExecuteNonQuery();

												//// Store the message in a format as if we'd generated it.
												//var newData = DatabaseAccess.readSkill(Convert.ToInt32(localId));
												//var ChangeMessage = new StaffSkillChange(skillRef, parentRef, newData);
												MessageTracking.SetLastGeneratedXml(message.GetMessageName(), message.EntityRef.Value, skill.ToXml());

												if (isNew & isValid)
												{
													BusRefTranslator.SetBusRef(EntityTranslationNames.Skill, idParameter.Value.ToString(), skillRef);
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
}
