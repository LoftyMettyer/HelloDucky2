using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Service.Services;
using Nexus.WebAPI.Controllers;
using System.Security.Claims;
using Nexus.Sql_Repository;
using Nexus.Common.Enums;

namespace Nexus.WebAPI.Tests.Controllers
{
    /// <summary>
    /// Unit test procedures for the designer services
    /// </summary>
    [TestClass]
    public class DesignerControllerTest
    {

        SqlDesignRepository _mockRepository;
        DesignerService _mockService;
        DesignerController _mockController;
        ClaimsIdentity _claims;

        [TestInitialize]
        public void TestInitialize()
        {
            _claims = new ClaimsIdentity();
            _claims.AddClaim(new Claim(ClaimTypes.Name, "testUser"));
            _claims.AddClaim(new Claim(ClaimTypes.NameIdentifier, "088C6A78-E14A-41B0-AD93-4FB7D3ADE96C"));

            _mockRepository = new SqlDesignRepository();
            _mockService = new DesignerService(_mockRepository);
            _mockController = new DesignerController(_mockService, _claims, "fr-fr");

        }

        [TestMethod]
        [TestCategory("Designer services")]
        public void AddTable_FailsForExistingTable()
        {
            Assert.Fail("Not yet implemented");
        }

        [TestMethod]
        [TestCategory("Designer services")]
        public void AddNewLookupTable_Success()
        {
            var result = _mockService.AddTable("newlookup");
            Assert.AreEqual(result, DesignStatus.Success);
            Assert.Fail("Not yet implemented");
        }

    }
}
