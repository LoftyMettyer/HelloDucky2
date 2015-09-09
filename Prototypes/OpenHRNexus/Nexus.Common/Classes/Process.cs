using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Common.Classes
{
    public class Process : BaseEntity
    {

        public List<ProcessStep> Steps { get; set; }

        public WebForm GetFirstStep {
            get
            {
                return new WebForm();
            }
        }
    }

}

