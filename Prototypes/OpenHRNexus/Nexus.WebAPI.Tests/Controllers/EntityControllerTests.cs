﻿using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Enums;
using Nexus.Common.Models;
using Nexus.Service.Services;
using Nexus.WebAPI.Controllers;
using Nexus.Sql_Repository;

namespace Nexus.WebAPI.Tests.Controllers {
	[TestClass()]
	public class EntityControllerTests {

		SqlDataRepository _mockRepository;
		EntityService _mockService;
		EntityController _mockController;

		[TestInitialize]
		public void TestInitialize() {
			_mockRepository = new SqlDataRepository();
			_mockService = new EntityService(_mockRepository);
			_mockController = new EntityController(_mockService);
		}

		[TestMethod()]
		public void GetEntitiesTest_IsValidModel() {
			var result = _mockController.GetEntities(EntityType.DataEntry);
			Assert.IsTrue(result is IEnumerable<EntityModel>);
		}
	}
}