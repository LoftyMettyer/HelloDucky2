using System;
using Nexus.Common.Interfaces;

namespace Nexus.Common.Classes.DataFilters 
{
    public class RangeFilter : IReportDataFilter
    {
        public string Condition { get; set; }

        public DateTime? EndRange { get; set; }

        public DateTime? StartRange { get; set; }

        public int RecordRange {get; set;}


    }
}
