using System.ComponentModel.DataAnnotations;

namespace Nexus.Sql_Repository.DatabaseClasses.Data
{
    public class DictionaryItem
    {
        public int Id { get; set; }

        [StringLength(50)]
        public string Key { get; set; }
        [StringLength(10)]
        public string Language { get; set; }
        public string Text { get; set; }
    }
}
