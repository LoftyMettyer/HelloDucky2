﻿using System;
using Dapper;
using System.Linq;
using System.Data.SqlClient;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using StructureMap.Attributes;
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

								var su = c.Query<Picture>(@"SELECT 
													CASE WHEN LEN(picture) > convert(int,400) THEN 'JPEG' ELSE 'EMPTY' END AS ImageType,
													CASE WHEN LEN(picture) < 400 THEN NULL ELSE SUBSTRING(picture,401,datalength(picture)-400) END AS Picture
													FROM fusion.staff
													WHERE StaffID = @StaffID",
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

                if (su != null)
                {
										su.costCenter = su.costCenter == "" ? null : su.costCenter;
                    su.effectiveFromSpecified = true;
                    su.effectiveToSpecified = true;

                    return su;
                }
            }
            return null;
        }

        public static Contact readContact(int localId)
        {

            string sQuery = string.Format("SELECT * FROM fusion.staffContact WHERE ID_Contact = {0}", localId);

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                var custDA = new SqlDataAdapter(sQuery, c);
                var custDS = new DataSet();
                custDA.Fill(custDS, "contact");

								var contact = new Contact();

                var pRow = custDS.Tables["contact"].Rows[0];

                contact.isRecordInactive = pRow["isrecordinactive"] == DBNull.Value ? false : (bool?)pRow["isrecordinactive"];
                contact.id_Staff = (int?) pRow["id_staff"];
                contact.title = pRow["title"].ToString() == "" ? null : pRow["title"].ToString();
                contact.forenames = pRow["Forenames"].ToString() == "" ? null : pRow["Forenames"].ToString();
                contact.surname = pRow["Surname"].ToString() == "" ? null : pRow["Surname"].ToString();
								contact.contactType = pRow["contactType"].ToString() == "" ? null : pRow["contactType"].ToString();
                contact.relationshipType = pRow["relationshipType"].ToString() == "" ? null : pRow["relationshipType"].ToString();
                contact.workMobile = pRow["workMobile"].ToString() == "" ? null : pRow["workMobile"].ToString();
                contact.personalMobile = pRow["personalMobile"].ToString() == "" ? null : pRow["personalMobile"].ToString();
                contact.workPhoneNumber = pRow["workPhoneNumber"].ToString() == "" ? null : pRow["workPhoneNumber"].ToString();
                contact.homePhoneNumber = pRow["homePhoneNumber"].ToString() == "" ? null: pRow["homePhoneNumber"].ToString();
                contact.email = pRow["email"].ToString() == "" ? null : pRow["email"].ToString();
                contact.notes = pRow["notes"].ToString() == "" ? null : pRow["notes"].ToString();
								contact.homeAddressSpecified = true;

                if (pRow["AddressLine1"].ToString() != ""
                    || pRow["AddressLine2"].ToString() != ""
                    || pRow["AddressLine3"].ToString() != ""
                    || pRow["AddressLine4"].ToString() != ""
                    || pRow["AddressLine5"].ToString() != ""
                    || pRow["postcode"].ToString() != "")
                {

                    contact.homeAddress = new Address
                    {
                        addressLine1 = pRow["AddressLine1"].ToString() == "" ? null : pRow["AddressLine1"].ToString(),
                        addressLine2 = pRow["AddressLine2"].ToString(),
                        addressLine3 = pRow["AddressLine3"].ToString(),
                        addressLine4 = pRow["AddressLine4"].ToString(),
                        addressLine5 = pRow["AddressLine5"].ToString(),
                        postCode = pRow["postcode"].ToString() == "" ? null : pRow["postcode"].ToString()
                    };

                }

                return contact;
            }
        }

        public static Skill readSkill(int localId)
        {

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                Skill su =
                    c.Query<Skill>(@"SELECT * from Fusion.staffSkill WHERE ID_Skill = @SkillID",
                                     new
                                     {
                                         SkillID = localId
                                     }
                        ).FirstOrDefault();

                if (su != null)
                {
                    su.reference = su.reference == "" ? null : su.reference;
                    su.outcome = su.outcome == "" ? null : su.outcome;
                    su.trainingStartSpecified = true;
                    su.trainingEndSpecified = true;
                    su.validFromSpecified = true;
                    su.validToSpecified = true;
                    su.didNotAttendSpecified = true;
                    return su;
                }
            }
            return null;
        }

        public static LegalDocument readDocument(int localId)
        {

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                var su = c.Query<LegalDocument>(@"SELECT * FROM Fusion.staffLegalDocument WHERE ID_Document = @DocumentID",
                        new
                        {
                            DocumentID = localId
                        }
                        ).FirstOrDefault();

                if (su != null)
                {
                    su.documentReference = su.documentReference == "" ? null : su.documentReference;
                    su.requestedBy = su.requestedBy == "" ? null : su.requestedBy;
                    su.acceptedBy = su.acceptedBy == "" ? null : su.acceptedBy;
                    su.acceptedDateFieldSpecified = true;
                    su.requestedDateFieldSpecified = true;
                    return su;
                }

                return null;
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
            var sQuery = string.Format("SELECT * FROM fusion.staff WHERE StaffID = {0}", localId);

            using (var c = new SqlConnection(connectionString))
            {
                c.Open();

                var custDA = new SqlDataAdapter(sQuery, c);
                var custDS = new DataSet();
                custDA.Fill(custDS, "staff");

                var su = new Staff {homeAddress = new Address()};

                var pRow = custDS.Tables["staff"].Rows[0];

                su.isRecordInactive = pRow["isrecordinactive"] == DBNull.Value ? false:  (bool?)pRow["isrecordinactive"];
                su.title = pRow["title"].ToString() == "" ? null : pRow["title"].ToString();
                su.forenames = pRow["Forenames"].ToString() == "" ? null : pRow["Forenames"].ToString();
                su.surname = pRow["Surname"].ToString() == "" ? null : pRow["Surname"].ToString();
                su.preferredName = pRow["preferredName"].ToString() == "" ? null : pRow["preferredName"].ToString();
                su.payrollNumber = pRow["payrollNumber"].ToString() == "" ? null : pRow["payrollNumber"].ToString();

                if (!DBNull.Value.Equals(custDS.Tables["staff"].Rows[0]["DOB"]))
                {
                    su.dob = Convert.ToDateTime(custDS.Tables["staff"].Rows[0]["DOB"].ToString());
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


                su.leavingDateSpecified = true;
                if (!DBNull.Value.Equals(pRow["leavingDate"]))
                {
                    su.leavingDate = Convert.ToDateTime(pRow["leavingDate"].ToString());
                }

                su.leavingReason = pRow["leavingreason"].ToString() == "" ? null : pRow["leavingreason"].ToString();
                su.companyName = pRow["CompanyName"].ToString() == "" ? null : pRow["CompanyName"].ToString();
                su.jobTitle = pRow["jobTitle"].ToString() == "" ? null : pRow["jobTitle"].ToString();
                su.managerRef = pRow["managerRef"].ToString() == "" ? null : pRow["managerRef"].ToString();

								if (pRow["AddressLine1"].ToString() != ""
										|| pRow["AddressLine2"].ToString() != ""
										|| pRow["AddressLine3"].ToString() != ""
										|| pRow["AddressLine4"].ToString() != ""
										|| pRow["AddressLine5"].ToString() != ""
										|| pRow["postcode"].ToString() != "")
								{
									su.homeAddress.addressLine1 = pRow["AddressLine1"].ToString() == "" ? " " : pRow["AddressLine1"].ToString();
									su.homeAddress.addressLine2 = pRow["AddressLine2"].ToString();
									su.homeAddress.addressLine3 = pRow["AddressLine3"].ToString();
									su.homeAddress.addressLine4 = pRow["AddressLine4"].ToString();
									su.homeAddress.addressLine5 = pRow["AddressLine5"].ToString();
									su.homeAddress.postCode = pRow["postcode"].ToString() == "" ? " " : pRow["postcode"].ToString();
								}
								else
								{
									su.homeAddress = null;
								}

								su.nationalInsuranceNumber = pRow["nationalInsuranceNumber"].ToString() == "" ? null : pRow["nationalInsuranceNumber"].ToString();

                return su;
            }
        }


    }
}
