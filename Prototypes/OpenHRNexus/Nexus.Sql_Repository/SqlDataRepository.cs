using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using Nexus.Common.Classes;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;

namespace Nexus.Sql_Repository
{
	public class SqlDataRepository : SqlRepositoryContext, IDataRepository, IEntityRepository
    {
		public IEnumerable<DynamicDataModel> GetData(int id)
		{
			var result = Data
				.Where(c => c.Id == id);

			return result.ToList();
		
		}

		public IEnumerable<DynamicDataModel> GetData()
		{
			return Data.ToList();
		}


        public WebForm GetWebForm(int id)
        {
            var webForm = WebForms.Where(w => w.id == id).First();

            // TODO - Need these 2 because the above is not loading on demand. I'm sure there's some linq that does this, but off the top of my head I don't know what it is.
            List<WebFormField> fields = WebFormFields.ToList();
            List<WebFormFieldOption> options = WebFormFieldOptions.ToList();

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

            var result = new WebFormModel
            {
                form_id = webForm.id.ToString(),
                form_name = webForm.Name,
                form_fields = webForm.Fields
            };

            //webForm.RemoveFieldsWithNoAccess();

            // Find all the tables that we need

//          IProcessRepository getSomeStuff; //= new WebForm();
          //  var blah = getSomeStuff.GetBaseTableInForm();




            //string tableInForm = webForm.GetBaseTableInForm();


            //List<string> relationsInForm = webForm.RelationsInForm();
            //List<string> columnsInForm = webForm.ColumnsInForm();
            //List<string> filtersInForm = webForm.FiltersInForm();
            //List<string> ordersInForm = webForm.OrdersInForm();


            //// Populate somehow?
            //string dynamicSQL = string.Format("SELECT {0} FROM {1} {2} {3} {4}"
            //    , string.Join(",", columnsInForm)
            //    , string.Join("", relationsInForm)
            //    , tableInForm
            //    , string.Join("", filtersInForm)
            //    , string.Join("", ordersInForm));

            // Dynamically build up a class to hoof this dynamic SQL into
            DynamicDataModel data = new DynamicDataModel();

//            var sql = string.Format("SELECT {0} FROM Personnel", string.Join(", ", result.form_fields));

            var sql = string.Format("SELECT id, Surname AS Column1, HolidayTaken AS Column2 FROM Personnel base ");

            // Append security filter
            sql += string.Format("INNER JOIN [User] u ON u.RecordId = base.Id WHERE u.UserID = '{0}'", userId);
              
          //  IEnumerable< DynamicDataModel> dummyData = Database.SqlQuery<string>(sql).ToList();

            var dummyData = Database.SqlQuery<DynamicDataModel>(sql).SingleOrDefault();

         //   var dummyDataSimple = Database.SqlQuery<string>(sql).ToList();

            //Data.Load()

            //    Data = Database.SqlQuery<string>(sql).ToList();


            foreach (WebFormField element in result.form_fields)
            {
                var column = Columns.Where(c => c.Id == element.field_columnid).First();

                // Security implemented here?
             //   sql


                if (element.field_columnid == 2)
                {
                    element.field_value = dummyData.Column1;                   
                }

                if (element.field_columnid == 14)
                {
                    element.field_value = dummyData.Column2.ToString();
                }


            }


            return result;

        }

        public BusinessProcess GetBusinessProcess(int Id)
        {
            return Processes.Where(p => p.Id == Id).First();
        }

        public virtual DbSet<DynamicDataModel> Data { get; set; }

        public virtual DbSet<BusinessProcess> Processes { get; set; }

        public virtual DbSet<WebForm> WebForms { get; set; }
        public virtual DbSet<WebFormField> WebFormFields { get; set; }
        public virtual DbSet<WebFormFieldOption> WebFormFieldOptions { get; set; }


        // Metadata for the dynamic objects
        public virtual DbSet<DynamicColumn> Columns { get; set; }


    }
}
