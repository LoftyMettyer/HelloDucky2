using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using Dapper;
using System.Data;
using System.Data.SqlClient;

namespace Fusion.Connector.OpenHR.DatabaseAccess
{
    public class StaffRecordDb 
    {
        public StaffRecordDb(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        private string ConnectionString { get; set; }

        public string MessageContext { get; set; }
        
        public int InsertData(string xmlMessage)
        {
            using (var c = new SqlConnection(ConnectionString))
            {
                c.Open();

                if (xmlMessage == "") return 0;

                XNamespace aw = "http://advancedcomputersoftware.com/xml/fusion/socialCare";
                XDocument loaded = XDocument.Load(new StringReader(xmlMessage));

                XElement rootNode;
                XElement nodeOfInterest;
                XName rootName;

                string EntityRef;
                string ParentRef;
               
                rootName = loaded.Root.Name;

//                nodeOfInterest = loaded.Element(rootName);
                rootName = loaded.Root.Name;
                rootNode = loaded.Element(rootName);

                nodeOfInterest = rootNode.Descendants().ElementAt(1);
                EntityRef = rootNode.FirstAttribute.NextAttribute.Value;
                ParentRef = rootNode.FirstAttribute.NextAttribute.NextAttribute.Value;

               


                if (MessageContext != null)
                {
                    c.Execute("fusion.pSetFusionContext", new
                    {
                        MessageType = MessageContext
                    },
                    commandType: CommandType.StoredProcedure);
                }

                // Update the db
                int? ID = (int?)c.Query<int?>(@"EXEC fusion.pSetDataForMessage @messagetype, @id, @xml, @parentguid",
                    new
                    {
                        messagetype = rootName.LocalName,
                        id = 0,
                        xml = nodeOfInterest.ToString(),
                        parentguid = ParentRef
                    }).First();

                return ID.Value;

            }
        }


        public void UpdateData(int id, string xmlMessage)
        {
            using (var c = new SqlConnection(ConnectionString))
            {
                c.Open();

                if (xmlMessage == "") return;

                XNamespace aw = "http://advancedcomputersoftware.com/xml/fusion/socialCare";
                XDocument loaded = XDocument.Load(new StringReader(xmlMessage));

                XElement rootNode;
                XElement nodeOfInterest;
                XName rootName;
                
                string EntityRef;
                string ParentRef;

                rootName = loaded.Root.Name;
                rootNode = loaded.Element(rootName);

                nodeOfInterest = rootNode.Descendants().ElementAt(1);
                EntityRef = rootNode.FirstAttribute.NextAttribute.Value;
                ParentRef = rootNode.FirstAttribute.NextAttribute.NextAttribute.Value;


                if (MessageContext != null)
                {
                    c.Execute("fusion.pSetFusionContext", new
                    {
                        MessageType = MessageContext
                    },
                    commandType: CommandType.StoredProcedure);
                }


                // Update the db
                c.Execute(@"EXEC fusion.pSetDataForMessage @messagetype, @id, @xml, @parentguid",
                    new
                    {
                        messagetype = rootName.LocalName,
                        id = id,
                        xml = nodeOfInterest.ToString(),
                        parentguid = ParentRef
                    });
                
            }
        }

        public StaffRecord ReadData(int id, string messageType)
        {
            using (var c = new SqlConnection(ConnectionString))
            {
                c.Open();

                StaffRecord su = c.Query<StaffRecord>(@"EXEC fusion.spGetDataForMessage @messagetype, @id, 1, 1, 1",
         new
         {
             MessageType = messageType,
             id = id
         }
         , buffered: false).First();

                return su;
            }
        }

    }
}
