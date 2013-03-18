//using System;
//using Fusion.Connector.OpenHR.DatabaseAccess;
//using Fusion.Connector.OpenHR.Configuration;
//using Fusion.Messages.General;
//using Fusion.Messages.SocialCare;
//using Fusion.Core.Sql;
//using Fusion.Core.Sql.OutboundBuilder;
//using StructureMap.Attributes;

//namespace Fusion.Connector.OpenHR.Messaging
//{
//    public class GenericMessageBuilder : IOutboundBuilder 
//    {

//      public GenericMessageBuilder() { }

//      //public GenericMessageBuilder(string connectionString, IBusRefTranslator busRefTranslator, IFusionConfiguration config)
//      //{
//      //    this.connString = connectionString;
//      //    this.refTranslator = busRefTranslator;
//      //    this.config = config;
//      //}



//        private readonly string connString;
//       // private readonly IBusRefTranslator refTranslator;
//       // private readonly IFusionConfiguration config;

//        [SetterProperty]
//        public IBusRefTranslator BusRefTranslator { get; set; }

//        [SetterProperty]
//        public IFusionConfiguration config { get; set; }


//        public FusionMessage Build(SendFusionMessageRequest source)
//        {
//            string messageType;
//            Type myType;
//            Guid parentRef = Guid.Empty;

//            string xml = "";
//            string messageName = source.MessageType;

//            if (source.MessageType.ToString() == "staffChange" | source.MessageType.ToString() == "staffContractChange")
//            {
//                messageType = source.MessageType + "Request";
//                myType = Type.GetType("Fusion.Messages.SocialCare." + messageType + ",Fusion.Messages.SocialCare");

//            }
//            else
//            {
//                messageType = source.MessageType + "Message";
//                myType = Type.GetType("Fusion.Messages.SocialCare." + messageType + ",Fusion.Messages.SocialCare");
//            }


//            var staffMember = new StaffRecordDb(System.Configuration.ConfigurationManager.ConnectionStrings["db"].ConnectionString);


//            var myData = staffMember.ReadData(Convert.ToInt32(source.LocalId), messageName);

//         //   Guid busRef = this.refTranslator.GetBusRef(messageName, source.LocalId);

//            Guid busRef = BusRefTranslator.GetBusRef(messageName, source.LocalId);


//            if (myData.ParentID != null)
//            {
//                parentRef = BusRefTranslator.GetBusRef("StaffChange", myData.ParentID);
//            }



//            if (myData != null)
//            {
//                if (myData.XMLCode != null)
//                {

//                    xml = String.Format(myData.XMLCode, busRef.ToString(), parentRef.ToString());

//                    //// Has message changed
//                    //if (xml == myData.XMLLastMessage)
//                    //{
//                    //    xml = null;
//                    //}
//                }

//            }

//            if (xml != null)
//            {
//                myType = Type.GetType("Fusion.Messages.SocialCare." + messageType + ", Fusion.Messages.SocialCare");
//                FusionMessage theMessage = (FusionMessage)Activator.CreateInstance(myType);

//                theMessage.CreatedUtc = source.TriggerDate;
//                theMessage.Id = Guid.NewGuid();
//                theMessage.Originator = config.ServiceName;
//                theMessage.EntityRef = busRef;
//                theMessage.Xml = xml;

//                return theMessage;

//               // return new PublishRequest()
//               // {
//               //     CreatedUtc = theMessage.CreatedUtc,
//               //     Id = theMessage.Id,
//               //     Originator = theMessage.Originator,
//               //     EntityRef = busRef,
//               //     Message = theMessage
//               //};

//            }
//            else
//            {
//                return null;
//            }

//        }

//    }
//}
