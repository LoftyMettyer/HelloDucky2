using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace Nexus.Sql_Repository.Tests
{
    [TestClass]
    public class ExtensionTests
    {
        [TestMethod]
        [TestCategory("Extensions")]
        public void Extension_StringFormatPlaceholder_HandlesMissingVariables()
        {
            var template = "SOME BASIC TEXT{missingcode} MORE TEXT {missingcode2} YET MORE TEXT {we_18_9} {we_18_10} FINAL BIT OF TEXT";

            var result = template.FormatPlaceholder(null);
            Assert.AreEqual(template, result, "Extension does not handle null variable list");

            result = template.FormatPlaceholder(new Dictionary<string, object>());
            Assert.AreEqual(template, result, "Extension does not handle empty variable list");

            var data = new Dictionary<string, object>()
                {
                    { "we_18_9", "insertedtext" },
                    { "we_18_10", DateTime.Now }
            };
            result = template.FormatPlaceholder(data);

            Assert.IsTrue(result.Contains("missingcode"), "Does not handle given key not present in the dictionary");
            Assert.IsFalse(result.Contains("{we_18_9}"), "Valid dictionary item not processed");

        }


    }
}
