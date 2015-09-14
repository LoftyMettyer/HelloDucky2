using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Interfaces.Repository;
using Nexus.Service.Services;
using Nexus.Sql_Repository;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using Nexus.WebAPI.Controllers;
using System.Collections.Generic;
using System.Security.Claims;

namespace Nexus.WebAPI.Tests.Controllers
{
    [TestClass]
    public class ProcessControllerTests
    {

        static IProcessRepository _mockRepository = new SqlProcessRepository();
        DataService _mockService = new DataService(_mockRepository);
        ProcessController _mockController;
        ClaimsIdentity _claims;

        [TestInitialize]
        public void TestInitialize()
        {
            _claims = new ClaimsIdentity();
            _claims.AddClaim(new Claim(ClaimTypes.Name, "testUser"));
            _claims.AddClaim(new Claim(ClaimTypes.NameIdentifier, "088C6A78-E14A-41B0-AD93-4FB7D3ADE96C"));

            _mockController = new ProcessController(_mockService, _claims, "fr-fr");

        }

        [TestMethod]
        public void GetPendingProcesses_ReturnsValidList()
        {
            var result = _mockController.GetPendingProcesses();
            Assert.IsInstanceOfType(result, typeof(IEnumerable<ProcessInFlow>));

        }
    }
}
