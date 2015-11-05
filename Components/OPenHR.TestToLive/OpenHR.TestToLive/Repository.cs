using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Runtime.Serialization;
using System.Runtime.InteropServices;
using OpenHRTestToLive.Interfaces;
using OpenHRTestToLive.Enums;
using System.Data.SqlClient;
using System.Data.Entity.Core.EntityClient;
using System.Data.Entity.Validation;
using System.Text;

namespace OpenHRTestToLive
{
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class Repository : IRepository
	{

        private string _connection { get; set; }

        public void Connection(string userName, string password, string databaseName, string serverName)
        {
            var connection = new SqlConnection();

            SqlConnectionStringBuilder sqlBuilder = new SqlConnectionStringBuilder();
            sqlBuilder.DataSource = serverName;
            sqlBuilder.InitialCatalog = databaseName;
            sqlBuilder.IntegratedSecurity = false;
            sqlBuilder.UserID = userName;
            sqlBuilder.Password = password;

            string providerString = sqlBuilder.ToString();

            // Initialize the EntityConnectionStringBuilder.
            EntityConnectionStringBuilder entityBuilder = new EntityConnectionStringBuilder();
            entityBuilder.Provider = "System.Data.SqlClient";
            entityBuilder.ProviderConnectionString = providerString;

            // Set the Metadata location.
            entityBuilder.Metadata = @"res://*/Model1.csdl|res://*/Model1.ssdl|res://*/Model1.msl";

            _connection = entityBuilder.ToString();
        }

        public string ExportDefinition(int Id, string fileName) {

            var liveDb = new npg_openhr8_2Entities(_connection);
            var output = new StringBuilder();
            var settings = new XmlWriterSettings();
                settings.CloseOutput = true;

            var copiedObjects = new T2LClass();
            ExtractAll(copiedObjects, liveDb, Id);

            //---------------------------------------------------------------------------------------------------------------------------------
            // Write all

            LogData("Saving {0}...", fileName);
            ConfirmSize(copiedObjects);
            DataContractSerializer AllWFSerializer = new DataContractSerializer(copiedObjects.GetType());

            XmlWriter WFWriter = XmlWriter.Create(File.CreateText(string.Format(fileName)),settings);
            //XmlWriter WFWriter = XmlWriter.Create(output);

            AllWFSerializer.WriteObject(WFWriter, copiedObjects);
            WFWriter.Flush();
            WFWriter.Close();

            LogData("Done", null);

            //return output.ToString();
            return WFWriter.ToString();
            
        }
        
        public RepositoryStatus ImportDefinitions(string inputFile)
        {
            var importObjects = new T2LClass();

            var liveDb = new npg_openhr8_2Entities(_connection);


            // Load the XML
            LogData("Reading {0}...", inputFile);
      DataContractSerializer AllWFSerializer = new DataContractSerializer(importObjects.GetType());
            XmlReader WFReader = XmlReader.Create(inputFile);
            importObjects = (T2LClass)AllWFSerializer.ReadObject(WFReader);
            ConfirmSize(importObjects);

            // Get the max existing WF id
            int MaxWFId = GetMaxWFId(liveDb);
            LogData("Max existing WF Id: {0}", MaxWFId);

            // Get the imported WF id
            int ImportedWFId = importObjects.AllWorkflows.First().id;

            //AllLinks = new List<ASRSysWorkflowLinks>();
            //AllElements = new List<ASRSysWorkflowElement>();
            //AllColumns = new List<ASRSysWorkflowElementColumn>();
            //AllValidations = new List<ASRSysWorkflowElementValidation>();
            //AllExpressions = new List<ASRSysExpression>();
            //AllComponents = new List<ASRSysExprComponent>();
            //AllItems = new List<ASRSysWorkflowElementItem>();
            //AllValues = new List<ASRSysWorkflowElementItemValue>();

            // If the imported ID clashes with the existing id range, fixup all imported id's
            if (ImportedWFId <= MaxWFId)
            {
                MaxWFId = BumpWorkflowIDs(importObjects, liveDb, MaxWFId);
            }

            // Disable all imported workflows
            foreach (var workflow in importObjects.AllWorkflows)
            {
                workflow.enabled = false;
            }


            // Assign the data lists back to the EF structures
            liveDb.ASRSysWorkflows.Add(importObjects.AllWorkflows.First());
            liveDb.ASRSysWorkflowLinks.AddRange(importObjects.AllLinks);
            liveDb.ASRSysWorkflowElementValidations.AddRange(importObjects.AllValidations);
            liveDb.ASRSysWorkflowElements.AddRange(importObjects.AllElements);
            liveDb.ASRSysWorkflowElementItemValues.AddRange(importObjects.AllValues);
            liveDb.ASRSysWorkflowElementItems.AddRange(importObjects.AllItems);
            liveDb.ASRSysWorkflowElementColumns.AddRange(importObjects.AllColumns);
            liveDb.ASRSysExpressions.AddRange(importObjects.AllExpressions);
            liveDb.ASRSysExprComponents.AddRange(importObjects.AllComponents);
            // And save


            try
            {
                liveDb.SaveChanges();
            }
            catch (DbEntityValidationException ex)
            {
                // Retrieve the error messages as a list of strings.
                var errorMessages = ex.EntityValidationErrors
                        .SelectMany(x => x.ValidationErrors)
                        .Select(x => x.ErrorMessage);

                // Join the list to a single string.
                var fullErrorMessage = string.Join("; ", errorMessages);

                // Combine the original exception message with the new one.
                var exceptionMessage = string.Concat(ex.Message, " The validation errors are: ", fullErrorMessage);

                // Throw a new DbEntityValidationException with the improved exception message.
                //throw new DbEntityValidationException(exceptionMessage, ex.EntityValidationErrors);
                return RepositoryStatus.Error;
            }

            return RepositoryStatus.DefinitionsImported;
        }


        //	Entity Hierarchy: --------------------------------------------------------------------------------------------------------------
        //	ASRSysWorkflows
        //		-	ASRSysWorkflowLinks
        //		-	ASRSysWorkflowElements
        //			-	ASRSysWorkflowElementColumn
        //			-	ASRSysWorkflowElementValidation
        //				- ASRSysExpression
        //					-	ASRSysExprComponent
        //			-	ASRSysWorkflowElementItem
        //				-	ASRSysWorkflowElementItemValue

        //=====================================================================================================================================
        // Class Definitions
        //=====================================================================================================================================


		//=====================================================================================================================================
		// Main
		//=====================================================================================================================================

		public static void Main(string[] args)
		{
			T2LClass t2l = new T2LClass();

			using (npg_openhr8_2Entities db = new npg_openhr8_2Entities(""))
			{

				LogData("Loading all WF headers...",null);

  			int WFCount = db.ASRSysWorkflows.ToList().Count();
				LogData("{0} Records loaded.", WFCount);

				foreach (ASRSysWorkflows WFRecord in db.ASRSysWorkflows)
				{
					LogData(string.Format("id: {0} \tname: {1}", WFRecord.id, WFRecord.name),null);
				}
				Console.Write("Please enter the WF id to copy: ");
				int WFkey = Convert.ToInt32( Console.ReadLine() );

				//XmlWriterSettings Settings = new XmlWriterSettings();
				//Settings.OmitXmlDeclaration = false;
				//Settings.NamespaceHandling = System.Xml.NamespaceHandling.OmitDuplicates;
				//Settings.NewLineOnAttributes = true;
				//XmlWriter WFWriter = XmlWriter.Create(File.CreateText("workflows.xml"),Settings);

				if (WFkey > 0)  // Extract WF definitions
				{
					ExtractAll(t2l, db, WFkey);

					//---------------------------------------------------------------------------------------------------------------------------------
					// Write all

					LogData("Saving allworkflow.xml...", null);
					ConfirmSize(t2l);
					DataContractSerializer AllWFSerializer = new DataContractSerializer(t2l.GetType());
					XmlWriter WFWriter = XmlWriter.Create(File.CreateText(string.Format("allworkflow.xml")));
					AllWFSerializer.WriteObject(WFWriter, t2l);
					WFWriter.Flush();
					WFWriter.Close();
					LogData("Done", null);
				}
				else if (WFkey == 0) // Load WF definition
				{
					// Load the XML
					LogData("Reading allworkflow.xml...", null);
					DataContractSerializer AllWFSerializer = new DataContractSerializer(t2l.GetType());
					XmlReader WFReader = XmlReader.Create("allworkflow.xml");
					t2l = (T2LClass)AllWFSerializer.ReadObject(WFReader);
					ConfirmSize(t2l);

					// Get the max existing WF id
					int MaxWFId = GetMaxWFId(db);
					LogData("Max existing WF Id: {0}", MaxWFId);

					// Get the imported WF id
					int ImportedWFId = t2l.AllWorkflows.First().id;

					//AllLinks = new List<ASRSysWorkflowLinks>();
					//AllElements = new List<ASRSysWorkflowElement>();
					//AllColumns = new List<ASRSysWorkflowElementColumn>();
					//AllValidations = new List<ASRSysWorkflowElementValidation>();
					//AllExpressions = new List<ASRSysExpression>();
					//AllComponents = new List<ASRSysExprComponent>();
					//AllItems = new List<ASRSysWorkflowElementItem>();
					//AllValues = new List<ASRSysWorkflowElementItemValue>();

					// If the imported ID clashes with the existing id range, fixup all imported id's
					if (ImportedWFId <= MaxWFId)
					{
						MaxWFId = BumpWorkflowIDs(t2l, db, MaxWFId);
					}

					// Assign the data lists back to the EF structures
					db.ASRSysWorkflows.Add(t2l.AllWorkflows.First());
					db.ASRSysWorkflowLinks.AddRange(t2l.AllLinks);
					db.ASRSysWorkflowElementValidations.AddRange(t2l.AllValidations);
					db.ASRSysWorkflowElements.AddRange(t2l.AllElements);
					db.ASRSysWorkflowElementItemValues.AddRange(t2l.AllValues);
					db.ASRSysWorkflowElementItems.AddRange(t2l.AllItems);
					db.ASRSysWorkflowElementColumns.AddRange(t2l.AllColumns);
					db.ASRSysExpressions.AddRange(t2l.AllExpressions);
					db.ASRSysExprComponents.AddRange(t2l.AllComponents);
					// And save
					db.SaveChanges();
				}
			}

			Console.ReadLine();
		}

		//=====================================================================================================================================
		// Utility
		//=====================================================================================================================================

		private static int BumpWorkflowIDs(T2LClass t2l, npg_openhr8_2Entities db, int MaxWFId)
		{
			// Update the descriptions
			string WFDescription = t2l.AllWorkflows.First().description;
			WFDescription = string.Concat("(T2L) ",WFDescription);
			t2l.AllWorkflows.First().description = WFDescription;

			// Bump the WorkflowID
			MaxWFId++;
			LogData("Bumping WF id's to {0}", MaxWFId);
			t2l.AllWorkflows.First().id = MaxWFId;

            foreach (ASRSysWorkflowElement item in t2l.AllElements) { item.WorkflowID = MaxWFId; }

            // Bump the Workflow Link ID's
            int MaxLinkID = db.ASRSysWorkflowLinks.Max(x => x.ID);
            MaxLinkID++;
            LogData("Bumping WF Link ID's to start at {0}", MaxLinkID);
            foreach (ASRSysWorkflowLinks item in t2l.AllLinks)
            {
                item.WorkflowID = MaxWFId;
                item.ID = MaxLinkID;
                MaxLinkID++;
            }


            int MaxElementItemID = db.ASRSysWorkflowElementItems.Max(x => x.ID);
            MaxElementItemID++;
            LogData("Bumping WF Element Item ID's to start at {0}", MaxElementItemID);
            foreach (ASRSysWorkflowElementItem child in t2l.AllItems)  // ID - Unique, ElementID - FK to WFElement.ID
            {
                // Bump the grandchildren
                int CurrentItemID = child.ID;
                child.ID = MaxElementItemID;
                foreach (ASRSysWorkflowElementItemValue grandchild in t2l.AllValues) // ItemID - FK to WFElementItem.ID
                {
                    if (grandchild.itemID == CurrentItemID) { grandchild.itemID = MaxElementItemID; }
                }
                MaxElementItemID++;
            }

            int MaxElementColumnID = db.ASRSysWorkflowElementColumns.Max(x => x.ID);
            MaxElementColumnID++;
            LogData("Bumping WF Element Columns ID's to start at {0}", MaxElementColumnID);
            foreach (ASRSysWorkflowElementColumn child in t2l.AllColumns)  // ID - Unique, ElementID - FK to WFElement.ID
            {
                child.ID = MaxElementColumnID;
                MaxElementColumnID++;
            }

            int MaxElementValidationID = db.ASRSysWorkflowElementValidations.Max(x => x.ID);
            MaxElementValidationID++;
            LogData("Bumping WF Element Validations ID's to start at {0}", MaxElementValidationID);
            foreach (ASRSysWorkflowElementValidation child in t2l.AllValidations)  // ID - Unique, ElementID - FK to WFElement.ID
            {
                child.ID = MaxElementValidationID;
                MaxElementValidationID++;
            }

            // Bump the workflow element ID's
            int MaxElementID = db.ASRSysWorkflowElements.Max(x => x.ID);
            MaxElementID++;

            LogData("Bumping WF Element ID's to start at {0}", MaxElementID);
            foreach (ASRSysWorkflowElement item in t2l.AllElements) 
			{
				// Deal with the child records first
				int CurrentElementID = item.ID;
				foreach (ASRSysWorkflowElementColumn child in t2l.AllColumns) { 
					if (child.ElementID == CurrentElementID) { child.ElementID = MaxElementID; } 
				}
				foreach (ASRSysWorkflowElementValidation child in t2l.AllValidations)
				{
					if (child.ElementID == CurrentElementID) { child.ElementID = MaxElementID; }
				}
				foreach (ASRSysWorkflowElementItem child in t2l.AllItems)  // ID - Unique, ElementID - FK to WFElement.ID
				{
                    if (child.ElementID == CurrentElementID) { child.ElementID = MaxElementID; }
				}

                foreach (ASRSysWorkflowLinks link in t2l.AllLinks)  // ID - Unique, ElementID - FK to WFElement.ID
                {
                    if (link.StartElementID == CurrentElementID) { link.StartElementID = MaxElementID; }
                    if (link.EndElementID == CurrentElementID) { link.EndElementID = MaxElementID; }
                }

                // Bump the parent
                item.ID = MaxElementID;
				MaxElementID++;
			}
            // Bump the Expression Component ID's
            int MaxExprComponentID = db.ASRSysExprComponents.Max(x => x.ComponentID);
            MaxExprComponentID++;
            LogData("Bumping WF Expression Component ID's to start at {0}", MaxExprComponentID);
            foreach (ASRSysExprComponent item in t2l.AllComponents)
            {
                item.ComponentID = MaxExprComponentID;
                MaxExprComponentID++;
            }

            // Bump the Expression ID's
            int MaxExprID = db.ASRSysExpressions.Max(x => x.ExprID);
            MaxExprID++;
            LogData("Bumping WF Expression ID's to start at {0}", MaxExprID);
            foreach (ASRSysExpression item in t2l.AllExpressions)
            {
                int CurrentExprID = item.ExprID;
                // Bump the PK
                item.ExprID = MaxExprID;
                item.UtilityID = MaxWFId;
                // Bunp the FK's
                foreach (ASRSysExprComponent child in t2l.AllComponents)
                {
                    if (child.ExprID == CurrentExprID) { child.ExprID = MaxExprID; }
                }
                foreach (ASRSysWorkflowElementValidation child in t2l.AllValidations)
                {
                    if (child.ExprID == CurrentExprID) { child.ExprID = MaxExprID; }
                }
                foreach (ASRSysWorkflowElementItem child in t2l.AllItems.Where(i => i.CalcID == CurrentExprID))
                {
                    child.CalcID = MaxExprID;
                }

                MaxExprID++;
            }

            return MaxWFId;
		}

		//-------------------------------------------------------------------------------------------------------------------------------------

		private static int GetMaxWFId(npg_openhr8_2Entities db)
		{
            int MaxId = db.ASRSysWorkflows.Max(x => x.id);
			return MaxId;
		}

		private static void ExtractAll(T2LClass t2l, npg_openhr8_2Entities db, int WFkey)
		{
			// Select single ASRSysWorkflows
			var SingleWFRecord = db.ASRSysWorkflows.Where(x => x.id == WFkey);  // int WFkey
            SingleWFRecord.First().description = SingleWFRecord.First().description.TrimEnd(' ');
            //WFCount = SingleWFRecord.ToList().Count();
            LogData("{0} WF Records loaded.", SingleWFRecord.ToList().Count());
			t2l.AllWorkflows.AddRange(SingleWFRecord);

			//DataContractSerializer WFSerializer = new DataContractSerializer(SingleWFRecord.GetType());
			//WFSerializer.WriteObject(WFWriter, SingleWFRecord.ToList());
			//WFWriter.Flush();
			//WFWriter.Close();

			// Select all child ASRSysWorkflowLinks
			t2l.AllLinks = db.ASRSysWorkflowLinks.Where(x => x.WorkflowID == WFkey).ToList();
			LogData("{0} WFLinks Records loaded.", t2l.AllLinks.Count());
			//DataContractSerializer WFLinksSerializer = new DataContractSerializer(t2l.AllLinks.GetType());
			//WFWriter = XmlWriter.Create(File.CreateText("workflowlinks.xml"));
			//WFLinksSerializer.WriteObject(WFWriter, t2l.AllLinks.ToList());
			//WFWriter.Flush();
			//WFWriter.Close();

			// Select all child ASRSysWorkflowElement
			t2l.AllElements = db.ASRSysWorkflowElements.Where(x => x.WorkflowID == WFkey).ToList();
			LogData("{0} WFElements Records loaded.", t2l.AllElements.Count());
			//DataContractSerializer WFElementsSerializer = new DataContractSerializer(t2l.AllElements.GetType());
			//WFWriter = XmlWriter.Create(File.CreateText("workflowelements.xml"));
			//WFElementsSerializer.WriteObject(WFWriter, t2l.AllElements.ToList());
			//WFWriter.Flush();
			//WFWriter.Close();

			// For each WorkFlow Element, select all ElementColumn records
			int ElementId = 0;
			//List<ASRSysWorkflowElementColumn> AllColumns = new List<ASRSysWorkflowElementColumn>();
			foreach (ASRSysWorkflowElement Element in t2l.AllElements)
			{
				ElementId = Element.ID;
				LogData("Element ID: {0}", ElementId);
				var GChildWFElementColumn = db.ASRSysWorkflowElementColumns.Where(x => x.ElementID == ElementId);
				if (GChildWFElementColumn.Count() > 0)
				{
					t2l.AllColumns.AddRange(GChildWFElementColumn);
					LogData("{0} Element Column grandchild records found", GChildWFElementColumn.Count());
					LogData("Total: {0}", t2l.AllColumns.Count());
				}
            }
            //WFWriter = XmlWriter.Create(File.CreateText(string.Format("workflowelementcolumns.xml")));
            //DataContractSerializer WFElementColumnSerializer = new DataContractSerializer(t2l.AllColumns.GetType());
            //WFElementColumnSerializer. WriteObject(WFWriter, t2l.AllColumns.ToList());
            //WFWriter.Flush();
            //WFWriter.Close();

            // For each WorkFlow Element, select all ElementValidation records
            foreach (ASRSysWorkflowElement Element in t2l.AllElements)
			{
				ElementId = Element.ID;
				LogData("Element ID: {0}", ElementId);
				var GChildWFElementValidation = db.ASRSysWorkflowElementValidations.Where(x => x.ElementID == ElementId);
				if (GChildWFElementValidation.Count() > 0)
				{
					t2l.AllValidations.AddRange(GChildWFElementValidation);
					LogData("{0} Element Validation grandchild records found", GChildWFElementValidation.Count());
					LogData("Total: {0}", t2l.AllValidations.Count());
				}
				else
					LogData("No Element Validation grandchild records found", null);
			}
			//DataContractSerializer WFElementValidationSerializer = new DataContractSerializer(t2l.AllValidations.GetType());
			//WFWriter = XmlWriter.Create(File.CreateText(string.Format("workflowelementvalidation.xml")));
			//WFElementValidationSerializer.WriteObject(WFWriter, t2l.AllValidations.ToList());
			//WFWriter.Flush();
			//WFWriter.Close();

			// For each WorkFlow Element, select all ElementItem records
			foreach (ASRSysWorkflowElement Element in t2l.AllElements)
			{
				ElementId = Element.ID;
				LogData("Element ID: {0}", ElementId);
				var GChildWFElementItem = db.ASRSysWorkflowElementItems.Where(x => x.ElementID == ElementId);
				if (GChildWFElementItem.Count() > 0)
				{
					t2l.AllItems.AddRange(GChildWFElementItem);
					LogData("{0} Element Item grandchild records found", GChildWFElementItem.Count());
					LogData("Total: {0}", t2l.AllItems.Count());
				}
				else
					LogData("No Element Item grandchild records found", null);
			}
			//DataContractSerializer WFElementItemSerializer = new DataContractSerializer(t2l.AllItems.GetType());
			//WFWriter = XmlWriter.Create(File.CreateText(string.Format("workflowelementitem.xml")));
			//WFElementItemSerializer.WriteObject(WFWriter, t2l.AllItems.ToList());
			//WFWriter.Flush();
			//WFWriter.Close();

			// For each WorkFlow Element Item, select all ElementItemValue records
			int ElementItemId = 0;
			foreach (ASRSysWorkflowElementItem ElementItem in t2l.AllItems)
			{
				ElementItemId = ElementItem.ID;
				LogData("Element ID: {0}", ElementItemId);
				var GChildWFElementItemValue = db.ASRSysWorkflowElementItemValues.Where(x => x.itemID == ElementItemId);
				if (GChildWFElementItemValue.Count() > 0)
				{
					t2l.AllValues.AddRange(GChildWFElementItemValue);
					LogData("{0} Element Item Value great-grandchild records found", GChildWFElementItemValue.Count());
					LogData("Total: {0}", t2l.AllValues.Count());
				}
				else
					LogData("No Element Item Value great-grandchild records found", null);
            }
			//DataContractSerializer WFElementItemValueSerializer = new DataContractSerializer(t2l.AllValues.GetType());
			//WFWriter = XmlWriter.Create(File.CreateText(string.Format("workflowelementitemvalue.xml")));
			//WFElementItemValueSerializer.WriteObject(WFWriter, t2l.AllValues.ToList());
			//WFWriter.Flush();
			//WFWriter.Close();

			// Expression Records ---------------------------------------------------------------------------------------------
			int ExpressionId = 0;
			// - WF Element (DescriptionExprID)
			foreach (ASRSysWorkflowElement Element in t2l.AllElements)
			{
				if (Element.DescriptionExprID != null)
				{
					ExpressionId = (int)Element.DescriptionExprID;
					LogData("Expression ID: {0}", ExpressionId);
					var WFExpression = db.ASRSysExpressions.Where(x => x.ExprID == ExpressionId);
					if (WFExpression.Count() > 0)
					{
						t2l.AllExpressions.AddRange(WFExpression);
						LogData("{0} Expression grandchild records found in WF Element (DescriptionExprID)", WFExpression.Count());
						LogData("Total: {0}", t2l.AllExpressions.Count());
					}
					else
						LogData("No Expression grandchild records found in WF Element (DescriptionExprID)", null);
				}
			}
			// - WF Element (TrueFlowExprID)
			foreach (ASRSysWorkflowElement Element in t2l.AllElements)
			{
				if (Element.TrueFlowExprID != null)
				{
					ExpressionId = (int)Element.TrueFlowExprID;
					LogData("Expression ID: {0}", ExpressionId);
					var WFExpression = db.ASRSysExpressions.Where(x => x.ExprID == ExpressionId);
					if (WFExpression.Count() > 0)
					{
						t2l.AllExpressions.AddRange(WFExpression);
						LogData("{0} Expression grandchild records found in WF Element (TrueFlowExprID)", WFExpression.Count());
						LogData("Total: {0}", t2l.AllExpressions.Count());
					}
					else
						LogData("No Expression grandchild records found in WF Element (TrueFlowExprID)", null);
				}
			}
			// - WF Element Column
			foreach (ASRSysWorkflowElementColumn Column in t2l.AllColumns)
			{
				if (Column.CalcID != null)
				{
					ExpressionId = (int)Column.CalcID;
					LogData("Expression ID: {0}", ExpressionId);
					var WFExpression = db.ASRSysExpressions.Where(x => x.ExprID == ExpressionId);
					if (WFExpression.Count() > 0)
					{
						t2l.AllExpressions.AddRange(WFExpression);
						LogData("{0} Expression grandchild records found in WF Element Column", WFExpression.Count());
						LogData("Total: {0}", t2l.AllExpressions.Count());
					}
					else
						LogData("No Expression grandchild records found in WF Element Column", null);
				}
			}
			// - WF Element Item (CalcID)
			foreach (ASRSysWorkflowElementItem Item in t2l.AllItems)
			{
				if (Item.CalcID != null)
				{
					ExpressionId = (int)Item.CalcID;
					LogData("Expression ID: {0}", ExpressionId);
					var WFExpression = db.ASRSysExpressions.Where(x => x.ExprID == ExpressionId);
					if (WFExpression.Count() > 0)
					{
						t2l.AllExpressions.AddRange(WFExpression);
						LogData("{0} Expression grandchild records found in WF Element Item (CalcID)", WFExpression.Count());
						LogData("Total: {0}", t2l.AllExpressions.Count());
					}
					else
						LogData("No Expression grandchild records found in WF Element Item (CalcID)", null);
				}
			}
			// - WF Element Item (RecordFilterID)
			foreach (ASRSysWorkflowElementItem Item in t2l.AllItems)
			{
				if (Item.RecordFilterID != null)
				{
					ExpressionId = (int)Item.RecordFilterID;
					LogData("Expression ID: {0}", ExpressionId);
					var WFExpression = db.ASRSysExpressions.Where(x => x.ExprID == ExpressionId);
					if (WFExpression.Count() > 0)
					{
						t2l.AllExpressions.AddRange(WFExpression);
						LogData("{0} Expression grandchild records found in WF Element Item (RecordFilterID)", WFExpression.Count());
						LogData("Total: {0}", t2l.AllExpressions.Count());
					}
					else
						LogData("No Expression grandchild records found in WF Element Item (RecordFilterID)", null);
				}
			}
			// - WF Element Validation
			foreach (ASRSysWorkflowElementValidation Validation in t2l.AllValidations)
			{
				if (Validation.ExprID != null)
				{
					ExpressionId = Validation.ExprID;
					LogData("ExprID: {0}", ExpressionId);
					var WFExpression = db.ASRSysExpressions.Where(x => x.ExprID == ExpressionId);
					if (WFExpression.Count() > 0)
					{
						t2l.AllExpressions.AddRange(WFExpression);
						LogData("{0} Expression grandchild records found in WF Element Validation", WFExpression.Count());
						LogData("Total: {0}", t2l.AllExpressions.Count());
					}
					else
						LogData("No Expression grandchild records found in WF Element Validation", null);
				}
			}

			// Expression Components
			List<ASRSysExprComponent> AllExpressionComponents = new List<ASRSysExprComponent>();
			FindExpressionComponents(db, t2l.AllExpressions, t2l.AllComponents);

			// Recursive Expressions
			List<ASRSysExpression> NewExpressions = new List<ASRSysExpression>();
			List<ASRSysExprComponent> NewExpressionComponents = new List<ASRSysExprComponent>();

			FindRecursiveExpressions(db, t2l.AllComponents, NewExpressions);
			while (NewExpressions.Count > 0)
			{
				FindExpressionComponents(db, NewExpressions, NewExpressionComponents);
				t2l.AllExpressions.AddRange(NewExpressions);
				NewExpressions.Clear();
				FindRecursiveExpressions(db, NewExpressionComponents, NewExpressions);
				t2l.AllComponents.AddRange(NewExpressionComponents);
				NewExpressionComponents.Clear();
			}
		}

		//-------------------------------------------------------------------------------------------------------------------------------------

		static void LogData(string s, object args)
		{
			Console.WriteLine(s,args);
        }

		//-------------------------------------------------------------------------------------------------------------------------------------

		static void ConfirmSize(T2LClass t2l)
		{
			LogData("Workflows: \t{0}",t2l.AllWorkflows.Count());
			LogData("Links: \t\t{0}",t2l.AllLinks.Count());
			LogData("Elements: \t{0}",t2l.AllElements.Count());
			LogData("Columns: \t{0}",t2l.AllColumns.Count());
			LogData("Validations: \t{0}",t2l.AllValidations.Count());
			LogData("Expressions: \t{0}",t2l.AllExpressions.Count());
			LogData("Components: \t{0}",t2l.AllComponents.Count());
			LogData("Items: \t\t{0}",t2l.AllItems.Count());
			LogData("Values: \t{0}",t2l.AllValues.Count());
		}

		//-------------------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Find all calculations, filters and field selection filters referenced in a set of Expression Components. 
		/// Return these as a list
		/// Expression/Component recursive referencing stops once the list is empty
		/// </summary>
		/// <param name="db"></param>
		/// <param name="ExpressionComponents"></param>
		/// <param name="Expressions"></param>
		private static void FindRecursiveExpressions(npg_openhr8_2Entities db, List<ASRSysExprComponent> ExpressionComponents, List<ASRSysExpression> Expressions)
		{
			int ExpressionId = 0;
			foreach (ASRSysExprComponent ExprComponent in ExpressionComponents)
			{
				ExpressionId = 0;

				// Expressions can be referenced from 3 fields
				if (ExprComponent.CalculationID != null)
				{
					ExpressionId = (int)ExprComponent.CalculationID;
					Console.WriteLine("Calculation ID: {0}", ExpressionId);
				}
				if (ExprComponent.FieldSelectionFilter != null)
				{
					ExpressionId = (int)ExprComponent.FieldSelectionFilter;
					Console.WriteLine("FieldSelectionFilter: {0}", ExpressionId);
				}
				if (ExprComponent.FilterID != null)
				{
					ExpressionId = (int)ExprComponent.FilterID;
					Console.WriteLine("FilterID: {0}", ExpressionId);
				}
				// Any non-zero, non-null value indicates an expression reference
				if (ExpressionId > 0)
				{
					var WFExpression = db.ASRSysExpressions.Where(x => x.ExprID == ExpressionId);
					if (WFExpression.Count() > 0)
					{
						Expressions.AddRange(WFExpression);
						Console.WriteLine("{0} Recursive Expression records found ", WFExpression.Count());
						Console.WriteLine("Total: {0}", Expressions.Count());
					}
				}
			}
		}

		//-------------------------------------------------------------------------------------------------------------------------------------
		/// <summary>
		/// Find all expression components referenced by a list of expressions.
		/// Return the results as a list.
		/// </summary>
		/// <param name="db"></param>
		/// <param name="Expressions"></param>
		/// <param name="ExpressionComponents"></param>
		private static void FindExpressionComponents(npg_openhr8_2Entities db, List<ASRSysExpression> Expressions, List<ASRSysExprComponent> ExpressionComponents)
		{
			int ExpressionId = 0;
			foreach (ASRSysExpression Expression in Expressions)
			{
				ExpressionId = Expression.ExprID;
				Console.WriteLine("Expression ID: {0}", ExpressionId);
				var WFExpressionComponents = db.ASRSysExprComponents.Where(x => x.ExprID == ExpressionId);
				if (WFExpressionComponents.Count() > 0)
				{
					ExpressionComponents.AddRange(WFExpressionComponents);
					Console.WriteLine("{0} Expression Component grandchild records found in Expr Component", WFExpressionComponents.Count());
					Console.WriteLine("Total: {0}", ExpressionComponents.Count());
				}
			}
		}

        private bool ValidateXML(string inputData)
        {
            return true;
        }

    }
}
