using System;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    public class TimesheetPerContract
    {

        public string contractName { get; set; }

        public DateTime? timesheetDate { get; set; }

        public decimal plannedHours { get; set; }

        public decimal workedHours { get; set; }

        public decimal toilHoursAccrued { get; set; }

        public decimal holidayHoursTaken { get; set; }

        public decimal toilHoursTaken { get; set; }

    }
}
