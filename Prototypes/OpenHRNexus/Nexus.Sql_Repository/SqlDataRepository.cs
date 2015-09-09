using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using Nexus.Common.Classes;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Sql_Repository.DatabaseClasses.Structure;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System.Data.Entity.Validation;
using Nexus.Common.Interfaces;
using Nexus.Sql_Repository.Enums;
using System.Data.Entity.Infrastructure;

namespace Nexus.Sql_Repository
{
	public class SqlDataRepository : SqlRepositoryContext, IDataRepository, IEntityRepository
    {
        bool _ExecuteImmediate = true;

        private List<WebFormFieldOption> GetLookupData(int columnId, string language)
        {
            var lookupTable = (from cols in Columns
                               where cols.Id == columnId
                               select cols).First();

            var formFields = (from cols in Columns
                              where cols.TableId == lookupTable.Id
                              select cols).ToList();
            var factory = new DynamicClassFactory();
            //            var dynamicType = CreateType(factory, string.Format("Lookup{0}", lookupTable.Id), formFields);

            var dynamicSQL = string.Format("SELECT * FROM Lookups WHERE Language = '{0}' AND WebFormField_id = {1}", language, columnId);

            //var dynamicSQL = string.Format("SELECT {0} FROM {1}",
            //     string.Join(", ", formFields.Select(c => c.PhysicalNameWithNullCheck)),
            //     lookupTable.PhysicalName);

            var data = Database.SqlQuery<WebFormFieldOption>(dynamicSQL);

             return data.ToList();



            //          SELECT id, column36 AS[title], column34 AS value, 17 AS WebFormField_id FROM userdefined4 where column35 = 'en-GB';
            //            SELECT id, column39 AS[title], column37 AS value, 21 AS WebFormField_id FROM userdefined5 where column38 = 'en-GB';

            //            SELECT id, column39 AS[title], column37 AS value, 21 AS WebFormField_id FROM userdefined5 where column38 = 'en-GB';


//            var result = new List<WebFormFieldOption>();

//            foreach (var row in data)
//            {
//                WebFormFieldOption dataRow = new WebFormFieldOption();

//                //dataRow.id
//result.Add(dataRow);


//                foreach (WebFormField element in result.fields)
//                {
//                    var property = row.GetType().GetProperty("column" + element.columnid);

//                    var value = property.GetValue(row, null);
//                    element.value = value == null ? string.Empty : value.ToString();
//                }
//            }



//            return data;
        }


        public WebForm GetWebForm(int id, string language)
        {
            var webForm = WebForms.Where(w => w.id == id).FirstOrDefault();

            // TODO - Need these 2 because the above is not loading on demand. I'm sure there's some linq that does this, but off the top of my head I don't know what it is.
            List<WebFormField> fields = WebFormFields.OrderBy(f => f.sequence).ToList();
            List<WebFormFieldOption> options = WebFormFieldOptions.ToList();
            List<WebFormButton> buttons = WebFormButtons.ToList();

            List<WebFormFieldOption> columnOptions;



            //            SELECT id, column36 AS[title], column34 AS value, 17 AS WebFormField_id FROM userdefined4 where column35 = 'en-GB';
            //          SELECT id, column39 AS[title], column37 AS value, 21 AS WebFormField_id FROM userdefined5 where column38 = 'en-GB';


            // Get lookup values and translate
                        

            foreach (var lookup in webForm.Fields.Where(f => f.columnid == 25)) {
                lookup.options = GetLookupData(lookup.columnid, language);
            }

            foreach (var lookup in webForm.Fields.Where(f => f.columnid == 21))
            {
                lookup.options = GetLookupData(lookup.columnid,language);
            }


            return webForm;
        }

        public IEnumerable<EntityModel> GetEntities(EntityType type)
        {

            IEnumerable<EntityModel> entities = new List<EntityModel>();

            switch (type)
            {
                case EntityType.Process:
                    //entities = Processes.Select(p => new EntityModel(p.Id, p.Name));

                    entities = Processes.Select(p => new EntityModel() { Id = p.Id, Name = p.Name });

                    break;

            }

            return entities;

        }

        public WebFormModel PopulateFormWithData(WebForm webForm, Guid userId)
        {

            var webFormId = webForm.id;

            var result = new WebFormModel
            {
                id = webForm.id,
                stepid = Guid.NewGuid(),
                name = webForm.Name,
                fields = webForm.Fields,
                buttons = webForm.Buttons
            };


            // Build column list
            var formFields = (from cols in Columns
                join form in WebFormFields on cols.Id equals form.columnid
                where form.WebForm.id == webFormId
                orderby form.sequence
                select cols).ToList();

            // Build tables
            var formTables = (from cols in Columns
                join form in WebFormFields on cols.Id equals form.columnid
                join t in DynamicTables on cols.TableId equals t.Id
                where form.WebForm.id == webFormId
                select t).ToList();

            // filter in security here???

            //            var tables = "Personnel";
            var tables = formTables.FirstOrDefault().PhysicalName;

            // TODO - The Dynamic type builder is not handling nulls and so we are forcing not nulls at this point
            // This causes an error below when we loop around the datarow.
            // Modify the class type builder to handle better.


            // Build select string
            var dynamicSQL = string.Format("SELECT id, {0} FROM {1} base ",
                             string.Join(", ", formFields.Select(c => c.PhysicalNameWithNullCheck)),
                tables);

            //string.Join(", ", formFields.Select(c => "ISNULL([" + c.PhysicalName + "], '') AS " + c.PhysicalName)),

            // Append security filter
            dynamicSQL += string.Format("INNER JOIN [User] u ON u.RecordId = base.Id WHERE u.UserID = '{0}'", userId);


            var factory = new DynamicClassFactory();
            var dynamicType = CreateType(factory, "webForm", formFields);
            var data = Database.SqlQuery(dynamicType, dynamicSQL);

            
            foreach (var row in data)
            {                            
                foreach (WebFormField element in result.fields)
                {
                    var property = row.GetType().GetProperty("column" + element.columnid);

                    var value = property.GetValue(row, null);
                    element.value = value == null ? string.Empty : value.ToString();
                }
            }
 
            return result;

        }

        public WebFormModel PopulateFormWithNavigationControls(WebForm webForm, Guid userId)
        {

            // Do the data opulation bit
            var formButtons = (from butt in WebFormButtons
                               where butt.WebForm.id == webForm.id
                               select butt).ToList();

            var result = new WebFormModel
            {
                id = webForm.id,
                name = webForm.Name,
                fields = webForm.Fields,
                buttons = webForm.Buttons
            };



            return result;
        }

        public Process GetProcess(int Id)
        {
            return Processes.Where(p => p.Id == Id).FirstOrDefault();
        }

        public virtual DbSet<WebForm> WebForms { get; set; }
        public virtual DbSet<WebFormField> WebFormFields { get; set; }
        public virtual DbSet<WebFormButton> WebFormButtons { get; set; }
        public virtual DbSet<WebFormFieldOption> WebFormFieldOptions { get; set; }

        // Metadata for the dynamic objects

        public virtual DbSet<DynamicColumn> Columns { get; set; }

        public virtual DbSet<DynamicTable> DynamicTables { get; set; }


        public virtual DbSet<Process> Processes { get; set; }
        public virtual DbSet<ProcessStep> ProcessSteps { get; set; }


        public virtual DbSet<ProcessInFlow> ProcessInFlow { get; set; }

        public virtual DbSet<TransactionStatement> Statements { get; set; }


        private Type CreateType(DynamicClassFactory dcf, string name, ICollection<DynamicColumn> dynamicAttributes)
        {

            // Original that creates the column as a name
            //            var props = dynamicAttributes.ToDictionary(da => da.DynamicAttribute.Name, da => typeof(string));
            var props = dynamicAttributes.ToDictionary(da => "column" + da.Id, da => da.DynamicDataType);

            var t = dcf.CreateDynamicType<BaseDynamicEntity>(name, props);
            return t;
        }

        public ProcessStepResponse SaveStepForLater(Guid stepId, Guid userID, WebFormModel form)
        {
            var response = new ProcessStepResponse();
            //    var formData = new ProcessInFlowData() { fields = form.fields };

//            var blah = form.fields.ToList();
  //          var step = new ProcessInFlow() { Id = stepId, UserId = userID, Data = form.fields.ToList() };


       //     var fieldValues = form.fields.ToList<ProcessInFlowData>();


//            ProcessInFlow.Add(step);

            try
            {
       //         SaveChanges();

                response = new ProcessStepResponse()
                {
                    Status = ProcessStepStatus.Success,
                    Message = "Success",
                    FollowOnUrl = String.Empty
                };

            }
            catch (DbEntityValidationException e)
            {
                response = new ProcessStepResponse()
                {
                    Status = ProcessStepStatus.ServerError,
                    Message = e.Message,
                    FollowOnUrl = String.Empty
                };

            }

            return response;

        }

        public IProcessStep GetProcessStep(Guid stepId)
        {
            return new ProcessStepEmail();
        }

        public IProcessStep GetProcessNextStep(IProcessStep currentStep)
        {
            return new ProcessStepEmail();
        }

        private ProcessStepResponse ExecuteStatemenForUser(string dynamicSQL, Guid UserId, bool Immediate)
        {
            var response = new ProcessStepResponse();

            var transaction = new TransactionStatement() {
                Id = Guid.NewGuid(),
                Statement = dynamicSQL,
                UserID = UserId,
                Time = DateTime.Now };

            try
            {
                Statements.Add(transaction);
                Database.ExecuteSqlCommand(transaction.Statement);
                SaveChanges();

                response = new ProcessStepResponse()
                {
                    Status = ProcessStepStatus.Success,
                    Message = "Success",
                    FollowOnUrl = String.Empty
                };

            }
            catch (Exception e)
            {
                response = new ProcessStepResponse()
                {
                    Status = ProcessStepStatus.ServerError,
                    Message = e.Message,
                    FollowOnUrl = String.Empty
                };

            }

            return response;
        }

        public ProcessStepResponse CommitStep(Guid stepId, Guid userId, WebFormModel data)
        {

            // Is Step insert/update/delete?
            var storedType = StoredDataType.Insert;




            // Get form field values in table/column enumerator
//            List<KeyValuePair<DynamicColumn, string>> dataValues;
     //       var BLAH = data.fields.Select(f => f.columnid).Contains(1);

            var columnIds = data.fields.Select(f => f.columnid).ToList();

            //var vals = (from cols in Columns
            //              where columnIds.Contains(cols.Id)
            //              select cols, "hello").ToList();


            var columns = (from cols in Columns
                               where columnIds.Contains(cols.Id)
                               select cols).ToList();

            var tableId = columns.FirstOrDefault().TableId;
            var table = DynamicTables.Where(t => t.Id == tableId).FirstOrDefault();

            var dynamicSQL = string.Format("INSERT [{0}] ({1}) VALUES ({2});",
                table.PhysicalName,
                string.Join(", ", columns.Select(c => "[" + c.PhysicalName + "]")),
                string.Join(", ", data.fields.OrderBy(c => c.columnid).Select(c => "'" + c.value + "'")));

            var response = ExecuteStatemenForUser(dynamicSQL, userId, _ExecuteImmediate);

            return response;
         

        }

        public IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters)
        {
            var results = new List<CalendarEventModel>();

            results.Add(new CalendarEventModel()
            {
                start = DateTime.Now,
                end = DateTime.Now.AddDays(2)
            });

            return  results;

        }
    }
}
