using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    public partial class Contract
    {

        private string contractNameField;

        private string departmentField;

        private string primarySiteField;

        private decimal contractedHoursPerWeekField;

        private decimal maximumHoursPerWeekField;

        private System.DateTime? effectiveFromField;

        private bool effectiveFromFieldSpecified;

        private DateTime? effectiveToField;

        private bool effectiveToFieldSpecified;

        public string contractName
        {
            get
            {
                return this.contractNameField;
            }
            set
            {
                this.contractNameField = value;
            }
        }

        public string department
        {
            get
            {
                return this.departmentField;
            }
            set
            {
                this.departmentField = value;
            }
        }

        public string primarySite
        {
            get
            {
                return this.primarySiteField;
            }
            set
            {
                this.primarySiteField = value;
            }
        }

        public decimal contractedHoursPerWeek
        {
            get
            {
                return this.contractedHoursPerWeekField;
            }
            set
            {
                this.contractedHoursPerWeekField = value;
            }
        }

        public decimal maximumHoursPerWeek
        {
            get
            {
                return this.maximumHoursPerWeekField;
            }
            set
            {
                this.maximumHoursPerWeekField = value;
            }
        }

        public DateTime? effectiveFrom
        {
            get
            {
                return this.effectiveFromField;
            }
            set
            {
                this.effectiveFromField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool effectiveFromSpecified
        {
            get
            {
                return this.effectiveFromFieldSpecified;
            }
            set
            {
                this.effectiveFromFieldSpecified = value;
            }
        }

        public DateTime? effectiveTo
        {
            get
            {
                return this.effectiveToField;
            }
            set
            {
                this.effectiveToField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool effectiveToSpecified
        {
            get
            {
                return this.effectiveToFieldSpecified;
            }
            set
            {
                this.effectiveToFieldSpecified = value;
            }
        }
    }

}
