using System;
using Nexus.Common.Interfaces;
using System.Data.Entity;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System.Linq;

namespace Nexus.Sql_Repository
{
    public class SqlDictionaryRepository : SqlRepositoryContext, IDictionary
    {
        private string _language;

        public string GetTranslation(string key)
        {
            var word = Dictionary.Where(d => d.Key == key && d.Language == _language).FirstOrDefault();

            if (word == null) {
                return key;
            }
            return word.Text;

        }

        public void SetLanguage(string language)
        {
            _language = language;
        }

        public virtual DbSet<DictionaryItem> Dictionary { get; set; }

    }


}
