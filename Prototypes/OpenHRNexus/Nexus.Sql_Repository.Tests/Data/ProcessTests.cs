using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Enums;
using System.Collections.Generic;
using Nexus.Common.Models;
using System.Linq;
using Nexus.Common.Interfaces.Repository;

namespace Nexus.Sql_Repository.Tests.Data
{
    [TestClass]
    public class ProcessTests
    {
        SqlDataRepository _mockRepository;

        [TestMethod]
        public void Process_GetList_ReturnsListOfValidProcesses()
        {
            _mockRepository = new SqlDataRepository();

            var result = _mockRepository.GetEntities(EntityType.Process);
            Assert.IsTrue(result.ToList().Count > 0);

        }

    }
}
