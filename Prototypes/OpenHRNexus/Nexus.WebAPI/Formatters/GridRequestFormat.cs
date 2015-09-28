using System.Collections;
using System.Collections.Generic;

namespace Nexus.WebAPI.Formatters
{
    /// <summary>
    /// Serialization class for outputting dynamic data to a response
    /// </summary>
    public class GridRequestFormat
    {
        public int total { get; set; }
        public int page { get; set; }
        public int records { get; set; }
        public IEnumerable rows { get; set; }
        public IEnumerable<ColumnDefinitionFormat> colModel { get; set; }
    }
}