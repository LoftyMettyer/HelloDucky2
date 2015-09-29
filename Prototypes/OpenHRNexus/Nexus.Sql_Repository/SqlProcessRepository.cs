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
using System.Net.Mail;
using Nexus.Sql_Repository.DatabaseClasses;
using System.Collections;
using System.Threading.Tasks;
using Nexus.Common.Classes.DataFilters;
using System.Web.Script.Serialization;

namespace Nexus.Sql_Repository
{
	public class SqlProcessRepository : SqlRepositoryContext, IProcessRepository, IEntityRepository
    {
        bool _ExecuteImmediate = true;

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

        public ProcessFormElement PopulateFormWithData(ProcessFormElement webForm, Guid userId)
        {

            var webFormId = webForm.id;

            var colIds = webForm.Fields.Select(s => s.columnid);

            // Build column list
            var formFields = (from col in Columns
                            where colIds.Contains(col.Id)
                            select col).ToList();


            // Build tables
            //var formTables = (from cols in Columns
            //    join form in fieldsInForm on cols.Id equals form.columnid
            //    join t in DynamicTables on cols.TableId equals t.Id
            //    where form.WebForm.id == webFormId
            //    select t).ToList();


            var formTables = (from col in Columns
                              where colIds.Contains(col.Id)
                              join t in DynamicTables on col.TableId equals t.Id
                              select t)
                              .Distinct()
                              .ToList();

            //var formTables = (from col in Columns
            //                   join t in DynamicTables on col.TableId equals t.Id
            //                   where colIds.Contains(col.Id)
            //                   select col).ToList();

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
            var dynamicType = CreateType(factory, "webForm", formFields, null);
            var data = Database.SqlQuery(dynamicType, dynamicSQL);

            
            foreach (var row in data)
            {                            
                foreach (WebFormField element in webForm.Fields)
                {
                    var property = row.GetType().GetProperty("column" + element.columnid);

                    var value = property.GetValue(row, null);
                    element.value = value == null ? string.Empty : value.ToString();
                }
            }
 
            return webForm;

        }
        

        [Obsolete("Will pull from database and calculate where in the process chain?")]
        private string GetBodyTemplateForEmail(IProcessStep step)
        {

            return "<!DOCTYPE html> " +
            "<html lang='en'>" +
            "    <head>" +
            "        <meta charset='utf-8' />" +
            "    </head>" +
            "    <body>" +
            "        <p>" +
            "            <span style='color: #0094ff'>{0}</span> has requested a <span style='color:#0094ff'>{1}</span> holiday absence from <span style='color:#0094ff'>{2}</span> to <span style='color:#0094ff'>{3}.</span>" +
            "        </p>" +
            "        <p>" +
            "            Reason for absence: <span style='color: #0094ff'>{4}</span>" +
            "        </p>" +
            "        <p>" +
            "            Employee notes: <span style='color: #0094ff'>{5}</span>" +
            "        </p>" +
            "        <p>" +
            "            You can quickly approve or decline this absence request using the buttons below." +
            "        </p>" +
            "<div>" +
            "<!--[if mso]>" +
            "<style type='text/css'>" +
            ".bold {{font-weight: bold}}" +
            "</style>" +
            "  <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='{6}' style='height:33px;v-text-anchor:middle;width:77px;margin-right: 5px;' arcsize='10%' stroke='f' fillcolor='#5CB85C'>" +
            "    <w:anchorlock/>" +
            "    <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:normal;'>" +
            "      Approve" +
            "    </center>" +
            "  </v:roundrect>" +
            "  <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='http://www.EXAMPLE.com/' style='height:33px;v-text-anchor:middle;width:77px;margin-right: 5px;' arcsize='10%' stroke='f' fillcolor='#D9534F'>" +
            "    <w:anchorlock/>" +
            "    <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:normal;'>" +
            "      Decline" +
            "    </center>" +
            "  </v:roundrect>" +
            "  <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='http://www.EXAMPLE.com/' style='height:33px;v-text-anchor:middle;width:130px;margin-right: 5px;' arcsize='10%' stroke='f' fillcolor='#5BC0DE'>" +
            "    <w:anchorlock/>" +
            "    <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:normal;'>" +
            "      View the request" +
            "    </center>" +
            "  </v:roundrect>" +
            "  <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='http://www.EXAMPLE.com/' style='height:33px;v-text-anchor:middle;width:149px;margin-right: 5px;' arcsize='10%' stroke='f' fillcolor='#337AB7'>" +
            "    <w:anchorlock/>" +
            "    <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:normal;'>" +
            "      View team calendar" +
            "    </center>" +
            "  </v:roundrect>" +
            "  <![endif]-->" +
            "  <!--[if !mso]>" +
            "{6}" +
            "  <![endif]-->" +
            "</div>" +
            "    </body>" +
            "</html>";
        }

        public ProcessEmailTemplate GetEmailTemplate(int id)
        {
            return ProcessEmailTemplates
                .Include("FollowOnActions")
                .AsNoTracking()
                .Where(t => t.Id == id)
                .First();
        }

        public Process GetProcess(int Id)
        {
            return Processes
                .Include("Elements")
                .Include("Elements.WebForm")
                .Include("Elements.WebForm.Fields")
                .Include("Elements.WebForm.Fields.options")
                .Include("Elements.WebForm.Buttons")
                .Where(p => p.Id == Id)
                .AsNoTracking()
                .FirstOrDefault();
                
        }

        public Process GetProcessForStep(Guid step)
        {
            return GetProcess(2);
        }

        public virtual DbSet<ProcessFormElement> WebForms { get; set; }
        public virtual DbSet<WebFormField> WebFormFields { get; set; }

        public virtual DbSet<ProcessEmailTemplate> ProcessEmailTemplates { get; set; }

        //public virtual DbSet<WebFormButton> WebFormButtons { get; set; }
        //public virtual DbSet<WebFormFieldOption> WebFormFieldOptions { get; set; }

        // Metadata for the dynamic objects

        public virtual DbSet<DynamicColumn> Columns { get; set; }

        public virtual DbSet<DynamicTable> DynamicTables { get; set; }


        public virtual DbSet<Process> Processes { get; set; }
//        public virtual DbSet<ProcessElement> ProcessElements { get; set; }


        public virtual DbSet<ProcessInFlow> ProcessInFlow { get; set; }
     //   public virtual DbSet<ProcessInFlowData> ProcessInFlowData { get; set; }


        public virtual DbSet<TransactionStatement> Statements { get; set; }


        private Type CreateType(DynamicClassFactory dcf, string name, ICollection<DynamicColumn> dynamicAttributes, List<Type> interfaces)
        {

            // Original that creates the column as a name
            //            var props = dynamicAttributes.ToDictionary(da => da.DynamicAttribute.Name, da => typeof(string));
            var props = dynamicAttributes.ToDictionary(da => "column" + da.Id, da => da.DynamicDataType);

            var t = dcf.CreateDynamicType<BaseDynamicEntity>(name, props, interfaces);
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
            return new ProcessEmailTemplate();
        }

        public IProcessStep GetProcessNextStep(IProcessStep currentStep)
        {
            return new ProcessEmailTemplate();
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
                //SaveChanges();

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
            var elementIds = data.fields.Select(f => f.elementid).ToList();

            var columns = (from fields in WebFormFields
                             where elementIds.Contains(fields.elementid)
                             join cols in Columns on fields.columnid equals cols.Id
                             orderby fields.elementid
                             select cols).ToList();


            var tableId = columns.FirstOrDefault().TableId;
            var table = DynamicTables.Where(t => t.Id == tableId).FirstOrDefault();

            var dynamicSQL = string.Format("INSERT [{0}] ({1}) VALUES ({2});",
                table.PhysicalName,
                string.Join(", ", columns.Select(c => "[" + c.PhysicalName + "]")),
                string.Join(", ", data.fields.OrderBy(c => c.elementid).Select(c => "'" + c.value + "'")));

            var response = ExecuteStatemenForUser(dynamicSQL, userId, _ExecuteImmediate);

            return response;
        
        }

        public WebFormDataModel UpdateProcessWithUserVariables(Process process, WebFormDataModel formData, Guid userId)
        {
            if (formData == null) return null;

            var instance = new ProcessInFlow() {
                Id = Guid.NewGuid(),
                InitiationUserId = userId,
                InitiationDateTime = DateTime.Now,
                ProcessName = process.Name,
                Caption = process.Name
            };

            instance.StepData.Add (new ProcessInFlowData()
            {
                Id = Guid.NewGuid(),
                UserId = userId,
                StepDateTime = DateTime.Now,
                StepData = new JavaScriptSerializer().Serialize(formData)
            });

            ProcessInFlow.Add(instance);
            SaveChanges();

            // Merge all the latest variables

            return formData;

        }

        public IEnumerable<ProcessInFlow> GetProcesses(Guid userId)
        {
            var result = ProcessInFlow
                .Where(u => u.InitiationUserId == userId)
                .Include("StepData");
            return result;

        }

        public IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters)
        {

            var dynamicSQL = string.Format("SELECT * FROM GetCalendarData"); //  WHERE Language = '{0}' AND WebFormField_id = {1}", _language, columnId);

            var data = Database.SqlQuery<CalendarEventModel>(dynamicSQL);
            return data.ToList();

        }


        public Type GetDataDefinition(int dataSourceId)
        {

            var factory = new DynamicClassFactory();
            var interfaces = new List<Type>() { typeof(IDynamicData) };
            var formFields = (from col in Columns
                              where col.TableId == dataSourceId
                              select col).ToList();

            return CreateType(factory, string.Format("dynamicData{0}", dataSourceId), formFields, interfaces);

        }


        public async Task<IEnumerable> GetData(int dataSourceId, IEnumerable<IReportDataFilter> filters)
        {

            // Build column list
            var formFields = (from col in Columns
                              where col.TableId == dataSourceId
                              select col).ToList();

            var baseTable = DynamicTables.Where(t => t.Id == dataSourceId).First().PhysicalName;

            // filter in security here???

            // TODO - The Dynamic type builder is not handling nulls and so we are forcing not nulls at this point
            // This causes an error below when we loop around the datarow.
            // Modify the class type builder to handle better.

            var dynamicTop = "";
            foreach (var filter in filters)
            {
                if (filter.GetType() == typeof(RangeFilter))
                {
                    dynamicTop = " TOP " + filter.RecordRange.ToString() + " ";
                }
               

            }


            // Build select string
            var dynamicSQL = string.Format("SELECT {0} id, {1} FROM {2} base ",
                             dynamicTop,
                             string.Join(", ", formFields.Select(c => c.PhysicalNameWithNullCheck)), baseTable);

            // Append security filter
//            dynamicSQL += string.Format("INNER JOIN [User] u ON u.RecordId = base.Id WHERE u.UserID = '{0}'", userId);


            var factory = new DynamicClassFactory();
            var interfaces = new List<Type>() { typeof(IDynamicData) };
            var dynamicType = CreateType(factory, string.Format("dynamicData{0}", dataSourceId), formFields, interfaces);

            var data = Database.SqlQuery(dynamicType, dynamicSQL);

            var dynamicData = await GetDynamicData(dynamicType, dynamicSQL);

            return dynamicData;        
        }

        private async Task<IEnumerable> GetDynamicData(Type type, string sql)
        {
            var data = Database.SqlQuery(type, sql);

  //          Task query = data.ToListAsync();
//            query.Wait();

            return await Database.SqlQuery(type, sql).ToListAsync();
        }

    }
}
