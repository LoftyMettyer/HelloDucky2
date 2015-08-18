using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenHRNexus.Common.Enums;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.SQLServer;
using OpenHRNexus.Service.Services;
using System.Collections.Generic;

namespace OpenHRNexus.WebAPI.Controllers.Tests
{
    [TestClass()]
    public class EntityControllerTests
    {

        SqlDataRepository _mockRepository;
        EntityService _mockService;
        EntityController _mockController;

        [TestInitialize]
        public void TestInitialize()
        {
            _mockRepository = new SqlDataRepository();
            _mockService = new EntityService(_mockRepository);
            _mockController = new EntityController(_mockService);
        }

        [TestMethod()]
        public void GetEntitiesTest_IsValidModel()
        {
            var result = _mockController.GetEntities(EntityType.DataEntry);
            Assert.IsTrue(result is IEnumerable<EntityModel>);
        }
    }
}