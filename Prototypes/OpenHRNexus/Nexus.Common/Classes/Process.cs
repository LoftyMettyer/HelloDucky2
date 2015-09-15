using Nexus.Common.Models;
using OpenHRNexus.Common.Enums;
using System.Collections.Generic;
using System.Linq;

namespace Nexus.Common.Classes
{
    public class Process : BaseEntity
    {

        public List<ProcessElement> Elements { get; set; }

        /// <summary>
        /// Returns the first ProcessFormElement for this entity
        /// </summary>
        /// <returns></returns>
        public ProcessFormElement GetEntryPoint()
        {
            var first = Elements
                .Where(e => e.Type == ProcessElementType.WebForm)
                .OrderBy(s => s.Sequence);

            return first.First().WebForm;
        }

    }

}

