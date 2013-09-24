﻿using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Dapper;
using Fusion.Connector.OpenHR.Database;
using Fusion.Core;
using log4net;
using NServiceBus;
using Fusion.Messages.SocialCare;
using Fusion.Connector.OpenHR.MessageComponents;

namespace Fusion.Connector.OpenHR.MessageHandlers
{
    public class StaffPictureChangeMessageHandler : BaseMessageHandler, IHandleMessages<StaffPictureChangeMessage>
    {

        private readonly string _connectionString;

        public StaffPictureChangeMessageHandler()
        {
            _connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            Logger = LogManager.GetLogger(typeof (StaffPictureChangeMessageHandler));
        }

        public void Handle(StaffPictureChangeMessage message)
        {

            var shouldProcess = StartHandlingMessage(message);

            if (!shouldProcess) return;

            StaffPictureChange picture;
            using (var sr = new StringReader(message.Xml))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    var serializer = new XmlSerializer(typeof (StaffPictureChange));
                    picture = (StaffPictureChange) serializer.Deserialize(xr);

                }
            }

            var busRef = new Guid(message.EntityRef.ToString());
            string localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, busRef);

            using (var c = new SqlConnection(_connectionString))
            {
                c.Open();

                var cmd = new SqlCommand("fusion.pMessageUpdate_StaffPictureChange", c)
                    {
                        CommandType = CommandType.StoredProcedure
                    };

                SqlParameter idParameter = cmd.Parameters.Add(new SqlParameter("@ID", localId ?? (object) DBNull.Value));
                idParameter.SqlDbType = SqlDbType.Int;
                idParameter.Direction = ParameterDirection.InputOutput;

                cmd.Parameters.Add(new SqlParameter("@recordIsInactive", picture.data.recordStatus));

                SqlParameter pictureParameter = cmd.Parameters.Add(new SqlParameter("@picture", picture.data.pictureChange.picture ?? (object) DBNull.Value));
                pictureParameter.SqlDbType = SqlDbType.Binary;
                pictureParameter.Direction = ParameterDirection.Input;


                try
                {
                    c.Execute("fusion.pSetFusionContext", new {MessageType = message.GetMessageName()},
                              commandType: CommandType.StoredProcedure);
                    cmd.ExecuteNonQuery();

										// Store the message in a format as if we'd generated it.
										var newData = DatabaseAccess.readPicture(Convert.ToInt32(localId));
										var ChangeMessage = new StaffPictureChange(busRef, newData);
										MessageTracking.SetLastGeneratedXml(message.GetMessageName(), message.EntityRef.Value, ChangeMessage.ToXml());


                }
                catch (Exception e)
                {
                    Logger.ErrorFormat("Inbound message {0}/{1} - {2} failed database save with error",
                                       message.GetMessageName(), message.EntityRef, e.Message);
                }

            }
        }
    }
}
