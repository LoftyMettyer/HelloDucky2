using System;
using System.Collections.Generic;
using Nexus.Common.Interfaces;
using Nexus.Common.Models;

namespace Nexus.Mock_Repository
{
    public class MockDictionaryRepository : ITranslation
    {
        public string Language { get; set; }

        public List<WebFormFieldOption> GetLookupValues(int columnId)
        {
            return new List<WebFormFieldOption>();
        }

        public string GetTranslation(string key)
        {
            return string.Format("{0}{1}", key, Language);
        }

    }
}
