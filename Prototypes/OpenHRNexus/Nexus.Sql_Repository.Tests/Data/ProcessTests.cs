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
        IDataRepository _mockRepository;

        //[TestMethod]
        //public void Process_GetList_ReturnsListOfValidProcesses()
        //{
        //    _mockRepository = new SqlDataRepository();

        //    var result = _mockRepository.GetEntities(EntityType.Process);
        //    Assert.IsTrue(result.ToList().Count > 0);

        //}

        [TestMethod]
        [TestCategory("Process")]
        public void Process_GetEntryPoint_AlwaysReturnsValidWebForm()
        {

            _mockRepository = new SqlDataRepository();

            var process = _mockRepository.GetProcess(1);
            var firstStep = process.GetEntryPoint();
            Assert.IsNotNull(firstStep);
            Assert.IsInstanceOfType(firstStep, typeof(WebForm));

        }
    }
}
