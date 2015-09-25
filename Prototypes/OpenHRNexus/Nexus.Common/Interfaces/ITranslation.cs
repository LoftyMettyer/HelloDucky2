using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Common.Interfaces
{
    public interface ITranslation
    {
        string GetTranslation(string key);
        List<WebFormFieldOption> GetLookupValues(int columnId);
        string Language { get; set; }
    }
}
