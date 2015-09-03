using Nexus.Common.Enums;
using Nexus.Common.Models;
using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace Nexus.Sql_Repository.DatabaseClasses.Structure
{
    public class DynamicColumn: EntityModel
    {
        //[ForeignKey("TableId")]
        //public virtual DynamicTable Table { get; set; }

        public int TableId { get; set; }

        public string DisplayName { get; set; }
        public ColumnDataType DataType { get; set; }

        public Type DynamicDataType
        {
            get
            {
                switch (DataType)
                {
                    case ColumnDataType.Decimal:
                        return Type.GetType("System.Decimal");
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

        public string PhysicalName
        {
            get
            {
                return string.Format("column{0}", Id);
            }
        }

        [Obsolete("As soon as the CreateType function can create nullable types this will serve no purpose")]
        public string PhysicalNameWithNullCheck
        {
            get
            {
                string nullValue;

                switch (DataType)
                {
                    case ColumnDataType.Decimal:
                    case ColumnDataType.Integer:
                        nullValue = "0";
                        break;
                    case ColumnDataType.DateTime:
                        nullValue = "''";
                        break;
                    case ColumnDataType.Boolean:
                        nullValue = "0";
                        break;
                    default:
                        nullValue = "''";
                        break;
                }

                return string.Format("ISNULL([column{0}], {1}) AS [column{0}]", Id, nullValue);
            }
        }


    }
}
