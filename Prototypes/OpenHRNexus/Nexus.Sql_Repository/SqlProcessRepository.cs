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
            var dynamicType = CreateType(factory, "webForm", formFields);
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
            // only one template exists at the moment
            //            var bodyTemplate = ProcessEmailTemplates.Where(t => t.Id == 1).First().BodyTemplate;
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
            //"  <![if !mso]>" +
            //"  <table cellspacing='0' cellpadding='0'> <tr> " +
            //"  <td align='center' width='300' height='40' bgcolor='#d62828' style='-webkit-border-radius: 5px; -moz-border-radius: 5px; border-radius: 5px; color: #ffffff; display: block;'>" +
            //"    <a href='http://www.EXAMPLE.com/' style='font-size:16px; font-weight: bold; font-family:sans-serif; text-decoration: none; line-height:40px; width:100%; display:inline-block'>" +
            //"    <span style='color: #ffffff;'>" +
            //"      Button Text Here!" +
            //"    </span>" +
            //"    </a>" +
            //"  </td> " +
            //"  </tr> </table> " +
            //"  <![endif]>" +
            "  <!--[if !mso]>" +
            "        <span style='background: green; padding: 5px'><a style='text-decoration: none; color: white' href='{6}'>Approve</a></span>" +
            "        <span style='background: red; padding: 5px'><a style='text-decoration: none; color: white' href='{7}'>Decline</a></span>" +
            "        <span style='background: lightblue; padding: 5px'><a style='text-decoration: none; color: white' href='{7}'>View the request</a></span>" +
            "        <span style='background: blue; padding: 5px'><a style='text-decoration: none; color: white' href='{9}'>View team calendar</a></span>" +
            "  <![endif]-->" +
            "</div>" +
            "    </body>" +
            "</html>";
        }

        public MailMessage PopulateEmailWithData(IProcessStep step, Guid userId, string targetURL, string authenticationToken, EmailAddressCollection destinations)
        {
            var emailStep = (ProcessStepEmail)step;

            var bodyTemplate = GetBodyTemplateForEmail(step);

            var buttonCode = string.Format("{0}UI/home/postprocessstep?userid={1}&code={3}&purpose={2}"
                , targetURL, userId, step.Id
                , authenticationToken);

            object[] dataBlob = { "Debbie Avery"
                            , "two day"
                            , "19/09/2015"
                            , DateTime.Now.AddDays(3)
                            , "Holiday"
                            , "Sorry it's short notice!"
                            , buttonCode
                            , "http://www.bbc.co.uk"
                            , "http://www.bbc.co.uk"
                            , "http://www.bbc.co.uk"
                    };

            var body = EmailFunctions.FormatBody(bodyTemplate, dataBlob, "en-GB");

            var result = new MailMessage(destinations.From, destinations.To)
            {
                Body = body
            };
            return result;

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

        public Guid RecordProcessStepForUser(ProcessFormElement form, Guid userID)
        {
            if (form == null) return Guid.Empty;

            var stepId = Guid.NewGuid();

            var process = new ProcessInFlow() { Id = stepId, UserId = userID, WebFormId = form .id};

            ProcessInFlow.Add(process);
          //  SaveChanges();

            return stepId;

        }

        public IEnumerable<ProcessInFlow> GetProcesses(Guid userId)
        {
            var result = ProcessInFlow
                .Where(u => u.UserId == userId);
            return result;

        }

    }
}
