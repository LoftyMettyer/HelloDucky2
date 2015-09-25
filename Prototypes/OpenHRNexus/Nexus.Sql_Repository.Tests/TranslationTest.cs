using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Interfaces;

namespace Nexus.Sql_Repository.Tests
{
    [TestClass]
    public class TranslationTest
    {

        ITranslation _translation = new SqlDictionaryRepository();

        [TestMethod]
        [TestCategory("Translation")]
        public void Dictionary_GetLookupReturnsList()
        {
            _translation.Language = "fr-fr";
            var result = _translation.GetLookupValues(25);
            Assert.IsNotNull(result);
        }

        [TestMethod]
        [TestCategory("Translation")]
        public void Dictionary_GetLookupReturnsListForUnicodeLanguage()
        {
            _translation.Language = "hi";
            var result = _translation.GetLookupValues(25);
            Assert.IsNotNull(result);
        }

        [TestMethod]
        [TestCategory("Translation")]
        public void Dictionary_GetLookupReturnsEnglishAsDefault()
        {
            var result = _translation.GetLookupValues(25);
            Assert.IsNotNull(result);
        }


        [TestMethod]
        [TestCategory("Translation")]
        public void Dictionary_TranslationReturnsValidResponse()
        {
            var text = "First Name";

            _translation.Language = "fr-fr";
            var resultFrench = _translation.GetTranslation(text);

            _translation.Language = "de-de";
            var resultGerman = _translation.GetTranslation(text);

            Assert.AreNotEqual(text, resultFrench);
            Assert.AreNotEqual(text, resultGerman);

        }

    }
}
