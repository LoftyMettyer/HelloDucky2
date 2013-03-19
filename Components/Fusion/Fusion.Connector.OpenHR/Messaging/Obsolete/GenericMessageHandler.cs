//using System;
//using NServiceBus;
//using log4net;
//using StructureMap.Attributes;
//using Fusion.Connector.OpenHR.DatabaseAccess;
//using Fusion.Messages.General;
//using Fusion.Messages.SocialCare;
//using Fusion.Core;
//using Fusion.Core.Sql;

//namespace Fusion.Connector.OpenHR.Messaging
//{
//    public class GenericMessageHandler : BaseMessageHandler, IHandleMessages<FusionMessage>
//    //, IHandleMessages<StaffChangeMessage>
//    //, IHandleMessages<StaffContractChangeMessage>
//    //, IHandleMessages<StaffPictureChangeMessage> 
//    //, IHandleMessages<StaffSkillChangeMessage>
//    //, IHandleMessages<StaffLegalDocumentChangeMessage>
//    //, IHandleMessages< StaffTimesheetPerContractChangeMessage>
//    //, IHandleMessages<StaffHolidayBalanceRemainingChangeMessage>
//    //, IHandleMessages<StaffContactChangeMessage>
//    //, IHandleMessages<OutreachAreaChangeMessage>
//    //, IHandleMessages<PersonTitleChangeMessage>
//    //, IHandleMessages<PersonMaritalStatusChangeMessage>
//    //, IHandleMessages<PersonEthnicityChangeMessage>
//    //, IHandleMessages<PersonContactTypeChangeMessage>
//    //, IHandleMessages<StaffLeavingReasonChangeMessage> 
//    //, IHandleMessages<StaffDepartmentChangeMessage>
//    //, IHandleMessages<StaffTrainingOutcomeChangeMessage>        
//    {
//        [SetterProperty]
//        public IBusRefTranslator BusRefTranslator { get; set; }

//        //public void Handle(StaffChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffContractChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffPictureChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffSkillChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffLegalDocumentChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffTimesheetPerContractChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffHolidayBalanceRemainingChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffContactChangeMessage message) { BaseHandle(message); }
//        //public void Handle(OutreachAreaChangeMessage message) { BaseHandle(message); }
//        //public void Handle(PersonTitleChangeMessage message) { BaseHandle(message); }
//        //public void Handle(PersonMaritalStatusChangeMessage message) { BaseHandle(message); }
//        //public void Handle(PersonEthnicityChangeMessage message) { BaseHandle(message); }
//        //public void Handle(PersonContactTypeChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffLeavingReasonChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffDepartmentChangeMessage message) { BaseHandle(message); }
//        //public void Handle(StaffTrainingOutcomeChangeMessage message) { BaseHandle(message); }


//        public void Handle(FusionMessage message)
//        {
//            bool shouldProcess = base.StartHandlingMessage(message);

//            Logger.InfoFormat("Inbound Message {0} Processing {1}", message.GetMessageName(), message.EntityRef);

//            if (shouldProcess == false) { return; }

//            string localId = BusRefTranslator.GetLocalRef(message.GetMessageName(), message.EntityRef.Value);

//            var StaffMember = new StaffRecordDb(System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString);
//            StaffMember.MessageContext = message.GetMessageName();

//            if (localId == null)
//            {
//                int newId = StaffMember.InsertData(message.Xml);
//                BusRefTranslator.SetBusRef(message.GetMessageName(), newId.ToString(), message.EntityRef.Value);
//            }
//            else
//            {
//                StaffMember.UpdateData(Convert.ToInt32(localId), message.Xml);
//            }

//        }

//        static GenericMessageHandler()
//        {
//            Logger = LogManager.GetLogger(typeof(FusionMessage));
//        }
//    }
//}
