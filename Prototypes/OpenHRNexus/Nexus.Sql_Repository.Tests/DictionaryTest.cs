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
        public void Dictionary_WebFormField_ConvertsToLanguage()
        {

            _dictionary.SetLanguage("en-GB");
            var webForm = new WebFormField(_dictionary);

            var keyValue = "untranslated";

            webForm.title = keyValue;
            Assert.AreNotEqual(webForm.title, keyValue);

        }

        public void Dictionary_WebFormField_EmptyDictionaryReturnsOriginalValue()
        {
            var webForm = new WebFormField();
            var keyValue = "untranslated";

            webForm.title = keyValue;
            Assert.AreEqual(webForm.title, keyValue);

        }

    }
}
