using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.SQLServer
{
	public class SqlDataRepository : DbContext, IDataRepository
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


        public IEnumerable<WebFormFields> GetWebFormFields(int id)
        {

            List<WebFormFields> fields = new List<WebFormFields>
            {
                new WebFormFields
                {
                    field_id = 1,
                    field_title = "First Name",
                    field_type = "textfield",
                    field_value = "John",
                    field_required = true,
                    field_disabled = false
                },
                new WebFormFields
                {
                    field_id = 2,
                    field_title = "Last Name",
                    field_type = "textfield",
                    field_value = "Doe",
                    field_required = true,
                    field_disabled = false
                }
            };

            List<WebFormFieldOption> options = new List<WebFormFieldOption>
            {
                new WebFormFieldOption
                {
                    option_id = 1,
                    option_title = "Male",
                    option_value = 1
                },
                new WebFormFieldOption
                {
                    option_id = 2,
                    option_title = "Female",
                    option_value = 2
                }
            };

            //List<WebFormFieldOption> options2 = WebFormFieldOptions.ToList();


            fields.Add(new WebFormFields
            {
                field_id = 3,
                field_title = "Gender",
                field_type = "radio",
                field_value = "2",
                field_required = true,
                field_disabled = false,
                field_options = options
            });

            fields.Add(new WebFormFields
            {
                field_id = 4,
                field_title = "Email Address",
                field_type = "email",
                field_value = "test@example.com",
                field_required = true,
                field_disabled = false
            });

            fields.Add(new WebFormFields
            {
                field_id = 5,
                field_title = "Password",
                field_type = "password",
                field_value = "",
                field_required = true,
                field_disabled = false
            });

            fields.Add(new WebFormFields
            {
                field_id = 6,
                field_title = "Birth Date",
                field_type = "date",
                field_value = "17.09.1971",
                field_required = true,
                field_disabled = false
            });

            options = new List<WebFormFieldOption>
            {
                new WebFormFieldOption
                {
                    option_id = 1,
                    option_title = "--Please Select--",
                    option_value = 1
                },
                new WebFormFieldOption
                {
                    option_id = 2,
                    option_title = "Internet Explorer",
                    option_value = 2
                },
                new WebFormFieldOption
                {
                    option_id = 3,
                    option_title = "Google Chrome",
                    option_value = 3
                },
                new WebFormFieldOption
                {
                    option_id = 4,
                    option_title = "Mozilla Firefox",
                    option_value = 4
                }
            };

            fields.Add(new WebFormFields
            {
                field_id = 7,
                field_title = "Your browser",
                field_type = "dropdown",
                field_value = "2",
                field_required = false,
                field_disabled = false,
                field_options = options
            });

            fields.Add(new WebFormFields
            {
                field_id = 8,
                field_title = "Additional Comments",
                field_type = "textarea",
                field_value = "Please type here...",
                field_required = false,
                field_disabled = false
            });

            fields.Add(new WebFormFields
            {
                field_id = 9,
                field_title = "I accept the terms and conditions",
                field_type = "checkbox",
                field_value = "0",
                field_required = true,
                field_disabled = false
            });

            fields.Add(new WebFormFields
            {
                field_id = 10,
                field_title = "I have a secret",
                field_type = "hidden",
                field_value = "X",
                field_required = false,
                field_disabled = false
            });

            return fields;
        }


        public virtual DbSet<DynamicDataModel> Data { get; set; }
        public virtual DbSet<WebFormFields> WebFormFields { get; set; }
        public virtual DbSet<WebFormFieldOption> WebFormFieldOptions { get; set; }

    }
}
