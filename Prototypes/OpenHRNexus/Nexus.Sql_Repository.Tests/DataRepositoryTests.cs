using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Nexus.Sql_Repository.Tests
{
    [TestClass]
    public class DataRepositoryTests
    {

        [TestMethod]
        [TestCategory("Get Data")]
        public void Repository_GetDataDefinition_ReturnsType()
        {
            var _mockRepository = new SqlProcessRepository();
            var result = _mockRepository.GetDataDefinition(1);

            Assert.IsInstanceOfType(result, typeof(Type), "Type class is not returned");
            Assert.IsNotNull(result.GetFields(), "No fields returned");

        }

    }
}
