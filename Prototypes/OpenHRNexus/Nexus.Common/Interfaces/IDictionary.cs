using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Common.Interfaces
{
    public interface IDictionary
    {
        string GetTranslation(string key);
        List<WebFormFieldOption> GetLookupValues(int columnId);
        string Language { get; set; }
    }
}
