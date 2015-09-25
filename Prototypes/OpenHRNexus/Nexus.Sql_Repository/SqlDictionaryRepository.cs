using System;
using Nexus.Common.Interfaces;
using System.Data.Entity;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System.Linq;
using Nexus.Common.Models;
using System.Collections.Generic;
using Nexus.Sql_Repository.DatabaseClasses.Structure;

namespace Nexus.Sql_Repository
{
    public class SqlDictionaryRepository : SqlRepositoryContext, ITranslation
    {
        private string _language = "en-GB";

        public string GetTranslation(string key)
        {
            var word = Dictionary.Where(d => d.Key == key && d.Language == Language).FirstOrDefault();

            var blah = Dictionary.ToList();


            if (word == null) {
                return key;
            }
            return word.Text;

        }

        public string Language
        {
            get { return _language; }
            set
            {
                _language = value;
            }
        }

        public virtual DbSet<DictionaryItem> Dictionary { get; set; }
        public virtual DbSet<DynamicColumn> Columns { get; set; }

        List<WebFormFieldOption> ITranslation.GetLookupValues(int columnId)
        {
            var lookupTable = (from cols in Columns
                               where cols.Id == columnId
                               select cols).First();

            var formFields = (from cols in Columns
                              where cols.TableId == lookupTable.Id
                              select cols).ToList();
            var factory = new DynamicClassFactory();
            //            var dynamicType = CreateType(factory, string.Format("Lookup{0}", lookupTable.Id), formFields);

            var dynamicSQL = string.Format("SELECT * FROM Lookups WHERE Language = '{0}' AND WebFormField_id = {1}", _language, columnId);

            //var dynamicSQL = string.Format("SELECT {0} FROM {1}",
            //     string.Join(", ", formFields.Select(c => c.PhysicalNameWithNullCheck)),
            //     lookupTable.PhysicalName);

            var data = Database.SqlQuery<WebFormFieldOption>(dynamicSQL);

            return data.ToList();
           
        }

    }


}
