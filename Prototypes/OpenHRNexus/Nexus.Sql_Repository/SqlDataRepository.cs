﻿using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using Nexus.Common.Classes;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Sql_Repository.DatabaseClasses.Structure;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System.Diagnostics;
using System.Data.Entity.Validation;
using System.ComponentModel.DataAnnotations.Schema;

namespace Nexus.Sql_Repository
{
	public class SqlDataRepository : SqlRepositoryContext, IDataRepository, IEntityRepository
    {

        public WebForm GetWebForm(int id)
        {
            var webForm = WebForms.Where(w => w.id == id).FirstOrDefault();

            // TODO - Need these 2 because the above is not loading on demand. I'm sure there's some linq that does this, but off the top of my head I don't know what it is.
            List<WebFormField> fields = WebFormFields.ToList();
            List<WebFormFieldOption> options = WebFormFieldOptions.ToList();
            List<WebFormButton> buttons = WebFormButtons.ToList();

            return webForm;
        }

        public IEnumerable<EntityModel> GetEntities(EntityType? id)
        {

            List<EntityModel> entities = new List<EntityModel>();

            entities.Add(new EntityModel
            {
                Id = 1,
                Name = "Personnel"
            });

            entities.Add(new EntityModel
            {
                Id = 2,
                Name = "Holiday Request"
            });

            entities.Add(new EntityModel
            {
                Id = 3,
                Name = "My Bank Details"
            });

            return entities;

        }

        public WebFormModel PopulateFormWithData(WebForm webForm, Guid userId)
        {

            var webFormId = webForm.id;

            var result = new WebFormModel
            {
                id = webForm.id.ToString(),
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
            var tables = formTables.FirstOrDefault().Name;

            // Build select string
            var dynamicSQL = string.Format("SELECT id, {0} FROM {1} base ",
                string.Join(", ", formFields.Select(c => "ISNULL([" +  c.Name + "], '') AS column_" + c.Id)),
                tables);


            // Append security filter
            dynamicSQL += string.Format("INNER JOIN [User] u ON u.RecordId = base.Id WHERE u.UserID = '{0}'", userId);


            var factory = new DynamicClassFactory();
            var dynamicType = CreateType(factory, "webForm", formFields);
            var data = Database.SqlQuery(dynamicType, dynamicSQL);



            
            foreach (var row in data)
            {                            
                foreach (WebFormField element in result.fields)
                {
                    var property = row.GetType().GetProperty("column_" + element.columnid);

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
                id = webForm.id.ToString(),
                name = webForm.Name,
                fields = webForm.Fields,
                buttons = webForm.Buttons
            };



            return result;
        }

        public BusinessProcess GetBusinessProcess(int Id)
        {
            return Processes.Where(p => p.Id == Id).FirstOrDefault();
        }

        public virtual DbSet<BusinessProcess> Processes { get; set; }

        public virtual DbSet<WebForm> WebForms { get; set; }
        public virtual DbSet<WebFormField> WebFormFields { get; set; }
        public virtual DbSet<WebFormButton> WebFormButtons { get; set; }
        public virtual DbSet<WebFormFieldOption> WebFormFieldOptions { get; set; }

        // Metadata for the dynamic objects

        public virtual DbSet<DynamicColumn> Columns { get; set; }

        public virtual DbSet<DynamicTable> DynamicTables { get; set; }

        public virtual DbSet<ProcessInFlow> ProcessInFlow { get; set; }

        private Type CreateType(DynamicClassFactory dcf, string name, ICollection<DynamicColumn> dynamicAttributes)
        {

            // Original that creates the column as a name
            //            var props = dynamicAttributes.ToDictionary(da => da.DynamicAttribute.Name, da => typeof(string));
            var props = dynamicAttributes.ToDictionary(da => "column_" + da.Id, da => da.DynamicDataType);

            var t = dcf.CreateDynamicType<BaseDynamicEntity>(name, props);
            return t;
        }

        public BusinessProcessStepResponse SaveStepForLater(Guid stepId, Guid userID, WebFormModel form)
        {
            var response = new BusinessProcessStepResponse();
            //    var formData = new ProcessInFlowData() { fields = form.fields };

//            var blah = form.fields.ToList();
  //          var step = new ProcessInFlow() { Id = stepId, UserId = userID, Data = form.fields.ToList() };


       //     var fieldValues = form.fields.ToList<ProcessInFlowData>();


//            ProcessInFlow.Add(step);

            try
            {
       //         SaveChanges();

                response = new BusinessProcessStepResponse()
                {
                    Status = BusinessProcessStepStatus.Success,
                    Message = "Success",
                    FollowOnUrl = String.Empty
                };

            }
            catch (DbEntityValidationException e)
            {
                response = new BusinessProcessStepResponse()
                {
                    Status = BusinessProcessStepStatus.ServerError,
                    Message = e.Message,
                    FollowOnUrl = String.Empty
                };

            }

            return response;

        }
    }
}
