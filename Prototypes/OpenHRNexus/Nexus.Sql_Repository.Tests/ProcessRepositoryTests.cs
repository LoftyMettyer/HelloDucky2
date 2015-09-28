using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Classes;
using Nexus.Common.Interfaces.Repository;
using System;
using System.Net.Mail;

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
            var processStepEmail = new ProcessStepEmail();

            var template = _mockRepository.GetEmailTemplate(1);

            Assert.IsNotNull(template.Destinations, "Email destinations is null");
            Assert.IsNotNull(template.Destinations.To, "Email To is null");
            Assert.IsNotNull(template.Destinations.From, "Email From is null");
        }

        [TestMethod]
        [TestCategory("Email")]
        public void Repository_PopulateEmailWithData_Populates()
        {
            var userId = Guid.NewGuid();
            var stepId = Guid.NewGuid();
            var processStepEmail = new ProcessStepEmail();

            var template = _mockRepository.GetEmailTemplate(1);

            var message = _mockRepository.PopulateEmailWithData(processStepEmail, userId, template);
            Assert.IsNotNull(message, "Populated email is returning null data");

        }



    }
}
