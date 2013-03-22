using System.Data.SqlClient;
using System.Linq;
using Fusion.Core;
using log4net;

using NServiceBus;
using StructureMap.Attributes;
using System.IO;
using System;
using Fusion.Core.Sql;
using Fusion.Messages.SocialCare;
using Fusion.Connector.OpenHR.MessageComponents;
using System.Xml.Serialization;
using System.Xml;
using Dapper;
using System.Data;

namespace Fusion.Connector.OpenHR.MessageHandlers
{
    public class StaffChangeMessageHandler : BaseMessageHandler, IHandleMessages<StaffChangeMessage>
    {

        [SetterProperty]
        public IBusRefTranslator BusRefTranslator {get;set;}

        private readonly string _connectionString;

        public StaffChangeMessageHandler ()
        {
            _connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            Logger = LogManager.GetLogger(typeof(StaffChangeMessageHandler));
        }

        public void Handle(StaffChangeMessage message)
        {

            SqlParameter idParameter;
            SqlParameter parameter;
            StaffChange _staff;

            bool shouldProcess = base.StartHandlingMessage(message);
            bool isNew = true;
            bool isValid = true;

            // testing hack - remove!!!!
            //shouldProcess = true;


            if (shouldProcess == true)
            {

                using (StringReader sr = new StringReader(message.Xml))
                {
                    using (XmlTextReader xr = new XmlTextReader(sr))
                    {
                        XmlSerializer serializer = new XmlSerializer(typeof (StaffChange));
                        _staff = (StaffChange) serializer.Deserialize(xr);

                    }
                }


                Guid busRef = new Guid(message.EntityRef.ToString());
                string localId = BusRefTranslator.GetLocalRef(EntityTranslationNames.Staff, busRef);

                // testing hack
                //localId = null;

                isNew = (localId  == null);


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

                    cmd.Parameters.Add(new SqlParameter("@title", _staff.data.staff.title ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@forenames", _staff.data.staff.forenames ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@surname", _staff.data.staff.surname ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@preferredName", _staff.data.staff.preferredName ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@payrollNumber", _staff.data.staff.payrollNumber ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@DOB", _staff.data.staff.DOB ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@employeeType", _staff.data.staff.employeeType ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@workMobile", _staff.data.staff.workMobile ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@personalMobile", _staff.data.staff.personalMobile ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@workPhoneNumber", _staff.data.staff.workPhoneNumber ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@homePhoneNumber", _staff.data.staff.homePhoneNumber ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@email", _staff.data.staff.email ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@personalEmail", _staff.data.staff.personalEmail ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@gender", _staff.data.staff.gender.ToString()));
                    cmd.Parameters.Add(new SqlParameter("@startDate", _staff.data.staff.startDate ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@leavingDate", _staff.data.staff.leavingDate ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@leavingReason", _staff.data.staff.leavingReason ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@companyName", _staff.data.staff.companyName ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@jobTitle", _staff.data.staff.jobTitle ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@managerRef", _staff.data.staff.managerRef ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@addressLine1", _staff.data.staff.homeAddress.addressLine1 ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@addressLine2", _staff.data.staff.homeAddress.addressLine2 ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@addressLine3", _staff.data.staff.homeAddress.addressLine3 ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@addressLine4", _staff.data.staff.homeAddress.addressLine4 ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@addressLine5", _staff.data.staff.homeAddress.addressLine5 ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@postcode", _staff.data.staff.homeAddress.postCode ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("@nationalInsuranceNumber", _staff.data.staff.nationalInsuranceNumber ?? (object)DBNull.Value));


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
                    BusRefTranslator.SetBusRef(EntityTranslationNames.Staff, idParameter.Value.ToString(), busRef);                    
                }

            }
        }

        //public void updateDb(string sql, string messageType)
        //{

        //    using (SqlConnection c = new SqlConnection(_connectionString))
        //    {
        //        c.Open();

        //        c.Execute("fusion.pSetFusionContext", new
        //            {
        //                MessageType = messageType
        //            },
        //                  commandType: CommandType.StoredProcedure);

        //        c.Execute(sql);

        //    }
        //}

//        public int insertIntoDb(string sql, string messageType)
//        {

//            using (SqlConnection c = new SqlConnection(_connectionString))
//            {
//                c.Open();

//                    c.Execute("fusion.pSetFusionContext", new
//                    {
//                        MessageType = messageType
//                    },
//                    commandType: CommandType.StoredProcedure);

//                return (int)c.Query<decimal>(sql,null).First();

//            }
//        }

//        private string formatForSQL(object data)
//        {
//            if (data != null)
//            {
//                string type = data.GetType().ToString();

//                switch (type)
//                {
//                    case "System.String":
//                        return string.Format("'{0}'", data);

//                    case "System.DateTime":
//                        return Convert.ToDateTime(data).ToString("yyyy-MM-dd");

//                    default:
////                        Enum.GetUnderlyingType(data.GetType())
//                       return string.Format("'{0}'", data);

//                }

//            }

//        return "NULL";

//        }

    }

}


