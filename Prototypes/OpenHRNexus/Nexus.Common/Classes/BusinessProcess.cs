using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Common.Classes
{
    public class BusinessProcess : BaseEntity
    {

        IEnumerable<BusinessProcessStep> Steps { get; set; }

        public WebForm GetFirstStep { get; }
    }

}

