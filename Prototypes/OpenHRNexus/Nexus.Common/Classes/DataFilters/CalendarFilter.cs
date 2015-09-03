using Nexus.Common.Interfaces;
using System;

namespace Nexus.Common.Classes.DataFilters
{
    public class CalendarFilter : IReportDataFilter

    {
        public string Condition { get; set; }

        public DateTime? EndRange { get; set; }

        public DateTime? StartRange { get; set; }

    }
}
