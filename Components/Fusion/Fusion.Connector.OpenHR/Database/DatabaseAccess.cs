﻿using System;
using Dapper;
using System.Linq;
using System.Data.SqlClient;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using StructureMap.Attributes;
using System.Data.SqlClient;
using System.Data;

namespace Fusion.Connector.OpenHR.Database
{
    public static class DatabaseAccess
    {

        [SetterProperty]
        public static IFusionConfiguration config { get; set; }

        private static string connectionString {get; set;}

        static DatabaseAccess()
        {
            connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString;
        }

        public static Picture readPicture(int localId)
        {

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                Picture su =
                    c.Query<Picture>(@"SELECT 'JPEG' AS ImageType, picture from Fusion.staff where StaffID = @StaffID",
                                     new
                                         {
                                             StaffID = localId
                                         }
                        ).FirstOrDefault();

                return su;
            }

        }

        public static Contract readContract(int localId)
        {

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                Contract su =
                    c.Query<Contract>(@"SELECT * from Fusion.staffContract WHERE ID_Contract = @ContractID",
                                     new
                                     {
                                         ContractID = localId
                                     }
                        ).FirstOrDefault();

                return su;

            }

        }

        public static Contact readContact(int localId)
        {

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                Contact su =
                    c.Query<Contact>(@"SELECT * from Fusion.staffContact WHERE ID_Contact = @ContactID",
                                     new
                                         {
                                             ContactID = localId
                                         }
                        ).FirstOrDefault();

                return su;

            }

        }

        public static Skill readSkill(int localId)
        {

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                Skill su =
                    c.Query<Skill>(@"SELECT * from Fusion.staffSkillChange WHERE ID_Skill = @SkillID",
                                     new
                                     {
                                         SkillID = localId
                                     }
                        ).FirstOrDefault();

                return su;

            }

        }

        public static LegalDocument readDocument(int localId)
        {

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                LegalDocument su = c.Query<LegalDocument>(@"SELECT * FROM Fusion.staffLegalDocument WHERE ID_Document = @DocumentID",
                        new
                        {
                            DocumentID = localId
                        }
                        ).FirstOrDefault();

                return su;

            }

        }

        public static TimesheetPerContract readTimesheet(int localId)
        {

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                TimesheetPerContract su =
                    c.Query<TimesheetPerContract>(@"SELECT * from Fusion.staffTimesheet WHERE ID_Timesheet = @TimesheetID",
                                     new
                                     {
                                         TimesheetID = localId
                                     }
                        ).FirstOrDefault();

                return su;

            }

        }

        public static Staff readStaff(int localId)
        {
            string value;

            string sQuery = string.Format("SELECT * FROM fusion.staff WHERE StaffID = {0}", localId);

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();


                // This uses a technique with the Dapper library
                // original - semi working, has orrible problems with the homeAddress child node. Possible fix with research, but don't have the time :-(
                //staffChangeDataStaff su = c.Query<staffChangeDataStaff>(@"SELECT Forenames, Surname, AddressLine1 FROM fusion.staff WHERE StaffID = @StaffID",
                //    new { StaffID = LocalID }).FirstOrDefault();

                //                SqlCommand selectCMD = new SqlCommand(sQuery, c);
                SqlDataAdapter custDA = new SqlDataAdapter(sQuery, c);
                DataSet custDS = new DataSet();
                custDA.Fill(custDS, "staff");


                Staff su = new Staff();
                su.homeAddress = new Address();
                //foreach (DataRow pRow in custDS.Tables["staff"].Rows)
                //{
                //    su.forenames = pRow["Forenames"].ToString();
                //}
                DataRow pRow = custDS.Tables["staff"].Rows[0];

                su.title = pRow["title"].ToString() == "" ? null : pRow["title"].ToString();
                su.forenames = pRow["Forenames"].ToString() == "" ? null : pRow["Forenames"].ToString();
                su.surname = pRow["Surname"].ToString() == "" ? null : pRow["Surname"].ToString();
                su.preferredName = pRow["preferredName"].ToString() == "" ? null : pRow["preferredName"].ToString();
                su.payrollNumber = pRow["payrollNumber"].ToString() == "" ? null : pRow["payrollNumber"].ToString();

                if (!DBNull.Value.Equals(custDS.Tables["staff"].Rows[0]["DOB"]))
                {
                    su.DOB = Convert.ToDateTime(custDS.Tables["staff"].Rows[0]["DOB"].ToString());
                }

                su.employeeType = pRow["employeeType"].ToString() == "" ? null : pRow["employeeType"].ToString();
                su.workMobile = pRow["workMobile"].ToString() == "" ? null : pRow["workMobile"].ToString();
                su.personalMobile = pRow["personalMobile"].ToString() == "" ? null : pRow["personalMobile"].ToString();
                su.workPhoneNumber = pRow["workPhoneNumber"].ToString() == "" ? null : pRow["workPhoneNumber"].ToString();
                su.homePhoneNumber = pRow["homePhoneNumber"].ToString() == "" ? null : pRow["homePhoneNumber"].ToString();
                su.email = pRow["email"].ToString() == "" ? null : pRow["email"].ToString();
                su.personalEmail = pRow["personalEmail"].ToString() == "" ? null : pRow["personalEmail"].ToString();

                if (custDS.Tables["staff"].Rows[0]["gender"].ToString() != "")
                {
                    su.gender = (Gender)Enum.Parse(typeof(Gender), custDS.Tables["staff"].Rows[0]["gender"].ToString(), true);
                }

                if (!DBNull.Value.Equals(custDS.Tables["staff"].Rows[0]["startDate"]))
                {
                    su.startDate = Convert.ToDateTime(custDS.Tables["staff"].Rows[0]["startDate"].ToString());
                }

                if (!DBNull.Value.Equals(custDS.Tables["staff"].Rows[0]["leavingDate"]))
                {
                    su.leavingDate = Convert.ToDateTime(custDS.Tables["staff"].Rows[0]["leavingDate"].ToString());
                }

                su.leavingReason = pRow["leavingreason"].ToString() == "" ? null : pRow["leavingreason"].ToString();
                su.companyName = pRow["CompanyName"].ToString() == "" ? null : pRow["CompanyName"].ToString();
                su.jobTitle = pRow["jobTitle"].ToString() == "" ? null : pRow["jobTitle"].ToString();
                su.managerRef = pRow["managerRef"].ToString() == "" ? null : pRow["managerRef"].ToString();
                su.homeAddress.addressLine1 = pRow["AddressLine1"].ToString();
                su.homeAddress.addressLine2 = pRow["AddressLine2"].ToString();
                su.homeAddress.addressLine3 = pRow["AddressLine3"].ToString();
                su.homeAddress.addressLine4 = pRow["AddressLine4"].ToString();
                su.homeAddress.addressLine5 = pRow["AddressLine5"].ToString();
                su.homeAddress.postCode = pRow["postcode"].ToString();
                su.nationalInsuranceNumber = pRow["nationalInsuranceNumber"].ToString() == "" ? null : pRow["nationalInsuranceNumber"].ToString();

                return su;
            }
        }


    }
}
