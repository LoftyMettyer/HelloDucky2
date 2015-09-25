using Nexus.Common.Interfaces;
using System;

namespace Nexus.Common.Classes.DataFilters
{
    public class DateRangeFilter : IReportDataFilter

    {
        public string Condition { get; set; }

        public DateTime? EndRange { get; set; }

        public int RecordRange { get; set; }

        public DateTime? StartRange { get; set; }

    }
}
