using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Nexus.Common.Interfaces.Repository;

namespace Nexus.Sql_Repository.Tests
{
    [TestClass]
    public class EmailGenerationTests
    {
        static IProcessRepository _mockRepository = new SqlProcessRepository();
        static SqlDictionaryRepository _mockDictionary = new SqlDictionaryRepository();

        [TestMethod]
        public void Email_GenerateEmail_IsSuccessfull()
        {

            var variables = new Dictionary<string, object>()
                {
                    {"we_1_1", "John" },
                    { "we_2_2", "Smith" },
                    {"we_22_8", 4 }
                };

            var template = _mockRepository.GetEmailTemplate(1);

            template._translation = _mockDictionary;
            template.Variables = variables;
            template.ConvertVariablesToDisplayText();




        }
    }
}
