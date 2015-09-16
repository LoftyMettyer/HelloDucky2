using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Interfaces;
using Nexus.Common.Models;
using Nexus.Mock_Repository;

namespace Nexus.Sql_Repository.Tests
{
    [TestClass]
    public class DictionaryTest
    {
//        IDictionary _dictionary = new MockDictionaryRepository();
        IDictionary _dictionary = new SqlDictionaryRepository();

        [TestMethod]
        [TestCategory("Translation")]
        public void Dictionary_GetLookupReturnsList()
        {
            _dictionary.Language = "fr-fr";
            var result = _dictionary.GetLookupValues(25);
            Assert.IsNotNull(result);
        }

        [TestMethod]
        [TestCategory("Translation")]
        public void Dictionary_GetLookupReturnsListForUnicodeLanguage()
        {
            _dictionary.Language = "hi";
            var result = _dictionary.GetLookupValues(25);
            Assert.IsNotNull(result);
        }

        [TestMethod]
        [TestCategory("Translation")]
        public void Dictionary_GetLookupReturnsEnglishAsDefault()
        {
            var result = _dictionary.GetLookupValues(25);
            Assert.IsNotNull(result);
        }


        [TestMethod]
        [TestCategory("Translation")]
        public void Dictionary_TranslationReturnsValidResponse()
        {
            var text = "First Name";

            _dictionary.Language = "fr-fr";
            var resultFrench = _dictionary.GetTranslation(text);

            _dictionary.Language = "de-de";
            var resultGerman = _dictionary.GetTranslation(text);

            Assert.AreNotEqual(text, resultFrench);
            Assert.AreNotEqual(text, resultGerman);

        }

    }
}
