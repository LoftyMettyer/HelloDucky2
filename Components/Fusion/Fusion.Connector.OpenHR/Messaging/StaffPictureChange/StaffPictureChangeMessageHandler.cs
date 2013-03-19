using System;
using System.Data.SqlClient;
using StructureMap.Attributes;
using log4net;
using Fusion.Core.Sql;
using Fusion.Connector.OpenHR.MessageHandlers;
using NServiceBus;
using Fusion.Messages.SocialCare;
using Fusion.Connector.OpenHR.MessageComponents.Component;


namespace Fusion.Connector.OpenHR.Messaging.StaffPictureChange
{
    public class StaffPictureChangeMessageHandler
    {
        public class StaffChangeMessageHandler : BaseMessageHandler, IHandleMessages<StaffPictureChangeMessage>
        {

            [SetterProperty]
            public IBusRefTranslator BusRefTranslator { get; set; }

            private readonly string _connectionString;

            public StaffChangeMessageHandler()
            {
                _connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                Logger = LogManager.GetLogger(typeof(StaffChangeMessageHandler));
            }

            public void Handle(StaffPictureChangeMessage message)
            {

                //SqlParameter idParameter;
                //SqlParameter parameter;
                //Picture _picture;


                //bool shouldProcess = base.StartHandlingMessage(message);
                //bool isNew = true;
                //bool isValid = true;

                //// testing hack - remove!!!!
                ////shouldProcess = true;


                //if (shouldProcess == true)
                //{

                //    using (StringReader sr = new StringReader(message.Xml))
                //    {
                //        using (XmlTextReader xr = new XmlTextReader(sr))
                //        {
                //            XmlSerializer serializer = new XmlSerializer(typeof(staffChange));
                //            _picture = (staffChange)serializer.Deserialize(xr);

                //        }
                //    }


                //    Guid busRef = new Guid(message.EntityRef.ToString());
                //    string localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, busRef);

                //    // testing hack
                //    //localId = null;

                //    isNew = (localId == null);


                //    using (var c = new SqlConnection(_connectionString))
                //    {
                //        c.Open();

                //        var cmd = new SqlCommand("fusion.pMessageUpdate_StaffChange", c)
                //        {
                //            CommandType = CommandType.StoredProcedure
                //        };

                //        idParameter = cmd.Parameters.Add(new SqlParameter("@ID", localId ?? (object)DBNull.Value));
                //        idParameter.SqlDbType = SqlDbType.Int;
                //        idParameter.Direction = ParameterDirection.InputOutput;

                //        cmd.Parameters.Add(new SqlParameter("@title", _picture.data.staff.title ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@forenames", _picture.data.staff.forenames ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@surname", _picture.data.staff.surname ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@preferredName", _picture.data.staff.preferredName ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@payrollNumber", _picture.data.staff.payrollNumber ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@DOB", _picture.data.staff.DOB ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@employeeType", _picture.data.staff.employeeType));
                //        cmd.Parameters.Add(new SqlParameter("@workMobile", _picture.data.staff.workMobile ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@personalMobile", _picture.data.staff.personalMobile ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@workPhoneNumber", _picture.data.staff.workPhoneNumber ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@homePhoneNumber", _picture.data.staff.homePhoneNumber ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@email", _picture.data.staff.email ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@personalEmail", _picture.data.staff.personalEmail ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@gender", _picture.data.staff.gender));
                //        cmd.Parameters.Add(new SqlParameter("@startDate", _picture.data.staff.startDate ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@leavingDate", _picture.data.staff.leavingDate ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@leavingReason", _picture.data.staff.leavingReason ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@companyName", _picture.data.staff.companyName ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@jobTitle", _picture.data.staff.jobTitle ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@managerRef", _picture.data.staff.managerRef ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@addressLine1", _picture.data.staff.homeAddress.addressLine1 ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@addressLine2", _picture.data.staff.homeAddress.addressLine2 ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@addressLine3", _picture.data.staff.homeAddress.addressLine3 ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@addressLine4", _picture.data.staff.homeAddress.addressLine4 ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@addressLine5", _picture.data.staff.homeAddress.addressLine5 ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@postcode", _picture.data.staff.homeAddress.postCode ?? (object)DBNull.Value));
                //        cmd.Parameters.Add(new SqlParameter("@nationalInsuranceNumber", _picture.data.staff.nationalInsuranceNumber ?? (object)DBNull.Value));

                //        try
                //        {
                //            c.Execute("fusion.pSetFusionContext", new { MessageType = message.GetMessageName() }, commandType: CommandType.StoredProcedure);
                //            cmd.ExecuteNonQuery();
                //        }
                //        catch (Exception e)
                //        {
                //            Logger.ErrorFormat("Inbound message {0}/{1} - {2} failed database save with error", message.GetMessageName(), message.EntityRef, e.Message);
                //            isValid = false;
                //        }

                //    }

                //    if (isNew & isValid)
                //    {
                //        BusRefTranslator.SetBusRef(EntityTranslationNames.Staff, idParameter.ToString(), busRef);
                //    }

                //}
            }

        }



    }
}
