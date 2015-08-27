using OpenHRNexus.Common.Models;
using System.Collections.Generic;

namespace OpenHRNexus.Service.Classes
{
    public class BusinessProcess : EntityModel
    {
        IEnumerable<BusinessProcessStep> Steps { get; set; }
    }
}
