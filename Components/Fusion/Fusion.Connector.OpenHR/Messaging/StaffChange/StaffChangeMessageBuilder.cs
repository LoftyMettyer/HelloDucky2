using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Connector.OpenHR.MessageComponents.Data;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using Fusion.Core.Sql;
using Fusion.Core.Sql.OutboundBuilder;
using Fusion.Messages.General;
using Fusion.Messages.SocialCare;
using StructureMap.Attributes;
using Fusion.Connector.OpenHR.Database;

using System.Xml;
using System.Xml.Serialization;
using System.IO;
using Fusion.Connector.OpenHR.MessageComponents.Component;

namespace Fusion.Connector.OpenHR.OutboundBuilders
{

    public class StaffChangeMessageBuilder : IOutboundBuilder
    {

        public string connectionString { get; set; }

        [SetterProperty]
        public IBusRefTranslator refTranslator { get; set; }

        [SetterProperty]
        public IFusionConfiguration config { get; set; }

        private Type _myType;
        private string _messageType;

        public FusionMessage Build(SendFusionMessageRequest source)
        {
            var busRef = refTranslator.GetBusRef(EntityTranslationNames.Staff, source.LocalId);
            var staff = DatabaseAccess.readStaff(Convert.ToInt32(source.LocalId));
            
            var xml = "";

            var xsSubmit = new XmlSerializer(typeof(StaffChange));
            var subReq = new StaffChange();

            subReq.data = new StaffChangeData
                {
                    staff = staff,
                    recordStatus = staff.isRecordInactive == true ? RecordStatusStandard.Inactive : RecordStatusStandard.Active,
                    auditUserName = "OpenHR user"
                };

            subReq.staffRef = busRef.ToString();

            var sww = new StringWriter();
            var writer = XmlWriter.Create(sww);
            xsSubmit.Serialize(writer, subReq);
            xml = sww.ToString();

            _messageType = source.MessageType + "Request";
            _myType = Type.GetType("Fusion.Messages.SocialCare." + _messageType + ", Fusion.Messages.SocialCare");

            var theMessage = (StaffChangeRequest)Activator.CreateInstance(_myType);

            theMessage.Community = ConfigurationManager.AppSettings["Community"];


            theMessage.PrimaryEntityRef = busRef;
            theMessage.CreatedUtc = source.TriggerDate;
            theMessage.Id = Guid.NewGuid();
            theMessage.Originator = config.ServiceName;
            theMessage.EntityRef = busRef;
            theMessage.Xml = xml;

            return theMessage;

        }


//        public Staff ReadData(int LocalID)
//        {
//            connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString;
//            string sQuery = string.Format("SELECT * FROM fusion.staff WHERE StaffID = {0}", LocalID);

//            using (SqlConnection c = new SqlConnection(connectionString))
//            {
//                c.Open();


//                // This uses a technique with the Dapper library
//                // original - semi working, has orrible problems with the homeAddress child node. Possible fix with research, but don't have the time :-(
//                //staffChangeDataStaff su = c.Query<staffChangeDataStaff>(@"SELECT Forenames, Surname, AddressLine1 FROM fusion.staff WHERE StaffID = @StaffID",
//                //    new { StaffID = LocalID }).FirstOrDefault();

////                SqlCommand selectCMD = new SqlCommand(sQuery, c);
//                SqlDataAdapter custDA = new SqlDataAdapter(sQuery, c);
//                DataSet custDS = new DataSet();
//                custDA.Fill(custDS, "staff");


//                Staff su = new Staff();
//                su.homeAddress = new Address();
//                //foreach (DataRow pRow in custDS.Tables["staff"].Rows)
//                //{
//                //    su.forenames = pRow["Forenames"].ToString();
//                //}
//                DataRow pRow = custDS.Tables["staff"].Rows[0];

//                su.title = pRow["title"].ToString() == "" ? null : pRow["title"].ToString();
//                su.forenames = pRow["Forenames"].ToString() == "" ? null : pRow["Forenames"].ToString();
//                su.surname = pRow["Surname"].ToString() == "" ? null : pRow["Surname"].ToString();
//                su.preferredName = pRow["preferredName"].ToString() == "" ? null : pRow["preferredName"].ToString();
//                su.payrollNumber = pRow["payrollNumber"].ToString() == "" ? null : pRow["payrollNumber"].ToString();

//                if (!DBNull.Value.Equals(custDS.Tables["staff"].Rows[0]["DOB"]))
//                {
//                    su.DOB = Convert.ToDateTime(custDS.Tables["staff"].Rows[0]["DOB"].ToString());
//                }

//                if (custDS.Tables["staff"].Rows[0]["employeeType"].ToString() != "")
//                {
//                    su.employeeType = (EmployeeType)Enum.Parse(typeof (EmployeeType),custDS.Tables["staff"].Rows[0]["employeeType"].ToString(), true);
//                }

//                su.workMobile = pRow["workMobile"].ToString() == "" ? null : pRow["workMobile"].ToString();
//                su.personalMobile = pRow["personalMobile"].ToString() == "" ? null : pRow["personalMobile"].ToString();
//                su.workPhoneNumber = pRow["workPhoneNumber"].ToString() == "" ? null : pRow["workPhoneNumber"].ToString();
//                su.homePhoneNumber = pRow["homePhoneNumber"].ToString() == "" ? null : pRow["homePhoneNumber"].ToString();
//                su.email = pRow["email"].ToString() == "" ? null : pRow["email"].ToString();
//                su.personalEmail = pRow["personalEmail"].ToString() == "" ? null : pRow["personalEmail"].ToString();

//                if (custDS.Tables["staff"].Rows[0]["gender"].ToString() != "")
//                {
//                    su.gender =(Gender)Enum.Parse(typeof (Gender), custDS.Tables["staff"].Rows[0]["gender"].ToString(), true);
//                }

//                if (!DBNull.Value.Equals(custDS.Tables["staff"].Rows[0]["startDate"]))
//                {
//                    su.startDate = Convert.ToDateTime(custDS.Tables["staff"].Rows[0]["startDate"].ToString());
//                }

//                if (!DBNull.Value.Equals(custDS.Tables["staff"].Rows[0]["leavingDate"]))
//                {
//                    su.leavingDate = Convert.ToDateTime(custDS.Tables["staff"].Rows[0]["leavingDate"].ToString());
//                }

//                su.leavingReason = pRow["leavingreason"].ToString() == "" ? null : pRow["leavingreason"].ToString();
//                su.companyName = pRow["CompanyName"].ToString() == "" ? null : pRow["CompanyName"].ToString();
//                su.jobTitle = pRow["jobTitle"].ToString() == "" ? null : pRow["jobTitle"].ToString();
//                su.managerRef = pRow["managerRef"].ToString() == "" ? null : pRow["managerRef"].ToString();
//                su.homeAddress.addressLine1 = pRow["AddressLine1"].ToString();
//                su.homeAddress.addressLine2 = pRow["AddressLine2"].ToString();
//                su.homeAddress.addressLine3 = pRow["AddressLine3"].ToString();
//                su.homeAddress.addressLine4 = pRow["AddressLine4"].ToString();
//                su.homeAddress.addressLine5 = pRow["AddressLine5"].ToString();
//                su.homeAddress.postCode = pRow["postcode"].ToString();
//                su.nationalInsuranceNumber = pRow["nationalInsuranceNumber"].ToString() == "" ? null : pRow["nationalInsuranceNumber"].ToString();


//                return su;
//            }
//        }
    
    }
}
