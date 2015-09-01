using Nexus.Common.Enums;
using Nexus.Common.Models;
using System;

namespace Nexus.Sql_Repository.DatabaseClasses.Structure
{
    public class DynamicColumn: EntityModel
    {


      //  [Column("TableId")]
    //    [ForeignKey("TableId")]
    //    public DynamicTable Table { get; set; }

        public int TableId { get; set; }
        //public int AttributeId { get; set; }
        public string DisplayName { get; set; }
        public ColumnDataType DataType { get; set; }

     //   [Key]
    //    public int Idx { get; set; }

        //[ForeignKey("AttributeId")]
        //public virtual DynamicAttribute DynamicAttribute { get; set; }

        //[ForeignKey("TemplateId")]
        //public virtual DynamicTable DynamicTemplate { get; set; }

        public Type DynamicDataType
        {
            get
            {
                switch (DataType)
                {
                    case ColumnDataType.Integer:
                        return Type.GetType("System.Int32");
                    case ColumnDataType.DateTime:
                        return Type.GetType("System.DateTime");
                    case ColumnDataType.Boolean:
                        return Type.GetType("System.Boolean");
                    default:
                        return Type.GetType("System.String");
                }
            }
        }

    }
}
