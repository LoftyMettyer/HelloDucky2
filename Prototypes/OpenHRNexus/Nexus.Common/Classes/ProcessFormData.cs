using System;
using System.Collections.Generic;

namespace Nexus.Common.Classes
{
    public class ProcessFormData
    {
        public Guid stepId { get; set; }
        public IEnumerable<KeyValuePair<string, object>> Data { get; set; }

    }
}
