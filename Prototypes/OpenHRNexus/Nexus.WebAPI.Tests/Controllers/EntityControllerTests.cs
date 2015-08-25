using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Enums;
using Nexus.Common.Models;
using Nexus.Repository.SQLServer;
using Nexus.Service.Services;
using System.Collections.Generic;

namespace Nexus.WebAPI.Controllers.Tests
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