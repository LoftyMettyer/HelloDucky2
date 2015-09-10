using Nexus.Common.Interfaces;

namespace Nexus.Mock_Repository
{
    public class MockDictionaryRepository : IDictionary
    {
        private string _language = "en-GB";

        public void SetLanguage(string language)
        {
            _language = language;
        }

        public string GetTranslation(string key)
        {
            return string.Format("{0}{1}", key, _language);
        }

    }
}
