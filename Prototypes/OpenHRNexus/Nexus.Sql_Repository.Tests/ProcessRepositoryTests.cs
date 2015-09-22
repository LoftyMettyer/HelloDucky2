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

            var destinations = processStepEmail.GetEmailDestinations();

            Assert.IsNotNull(destinations, "Email destinations is null");
            Assert.IsNotNull(destinations.To, "Email To is null");
            Assert.IsNotNull(destinations.From, "Email From is null");
        }

        [TestMethod]
        [TestCategory("Email")]
        public void Repository_PopulateEmailWithData_Populates()
        {
            var userId = Guid.NewGuid();
            var stepId = Guid.NewGuid();
            var processStepEmail = new ProcessStepEmail();

            var destinations = processStepEmail.GetEmailDestinations();

            var message = _mockRepository.PopulateEmailWithData(processStepEmail, userId, "<<TARGETURL>>", "<<AUTHENTICATIONCODE>>", destinations);
            Assert.IsNotNull(message, "Populated email is returning null data");

        }



    }
}
