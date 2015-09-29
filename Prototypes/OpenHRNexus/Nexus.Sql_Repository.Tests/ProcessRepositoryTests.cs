using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Classes;
using Nexus.Common.Interfaces.Repository;

namespace Nexus.Sql_Repository.Tests
{
    [TestClass]
    public class ProcessRepositoryTests
    {

        static IProcessRepository _mockRepository = new SqlProcessRepository();

        [TestMethod]
        [TestCategory("Email")]
        public void Repository_EmailDestinationsIsPopulated()
        {

            var template = _mockRepository.GetEmailTemplate(1);

            Assert.IsNotNull(template.Destinations, "Email destinations is null");
            Assert.IsNotNull(template.Destinations.To, "Email To is null");
            Assert.IsNotNull(template.Destinations.From, "Email From is null");
        }

    }
}
