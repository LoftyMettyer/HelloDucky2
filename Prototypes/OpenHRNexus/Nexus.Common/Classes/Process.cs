using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Common.Classes
{
    public class Process : BaseEntity
    {

        public List<ProcessStep> Steps { get; set; }

    }

}

