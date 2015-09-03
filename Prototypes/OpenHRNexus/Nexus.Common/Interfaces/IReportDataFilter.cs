using System;

namespace Nexus.Common.Interfaces
{
    public interface IReportDataFilter
    {
        string Condition { get; set; }
        DateTime? StartRange { get; set; }
        DateTime? EndRange { get; set; }
    }
}
