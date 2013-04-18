using System;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    public class TimesheetPerContract
    {

        public string contractName { get; set; }

        [XmlElementAttribute(DataType = "date")]
        public DateTime? timesheetDate { get; set; }

        public decimal? plannedHours { get; set; }

        public decimal? workedHours { get; set; }

        public decimal? toilHoursAccrued { get; set; }

        public decimal? holidayHoursTaken { get; set; }

        public decimal? toilHoursTaken { get; set; }

        [XmlIgnoreAttribute]
        public int? id_Staff { get; set; }

        [XmlIgnoreAttribute]
        public bool? isRecordInactive { get; set; }

    }
}
