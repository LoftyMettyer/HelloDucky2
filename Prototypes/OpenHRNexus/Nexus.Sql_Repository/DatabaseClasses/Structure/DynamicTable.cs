using Nexus.Common.Enums;
using Nexus.Common.Models;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace Nexus.Sql_Repository.DatabaseClasses.Structure
{
    public class DynamicTable : EntityModel
    {

        public TableType Type { get; set; }
        public string Description { get; set; }

        [ForeignKey("TableId")]
        public virtual ICollection<DynamicColumn> Columns { get; set; }

        public string PhysicalName
        {
            get
            {
                return string.Format("UserDefined{0}", Id);
            }
        }
    }
}
