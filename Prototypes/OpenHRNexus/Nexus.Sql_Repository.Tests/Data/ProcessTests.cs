using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Enums;
using System.Collections.Generic;
using Nexus.Common.Models;
using System.Linq;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Classes;

namespace Nexus.Sql_Repository.Tests.Data
{
    [TestClass]
    public class ProcessTests
    {
        IProcessRepository _mockRepository;

        //[TestMethod]
        //public void Process_GetList_ReturnsListOfValidProcesses()
        //{
        //    _mockRepository = new SqlProcessRepository();

        //    var result = _mockRepository.GetEntities(EntityType.Process);
        //    Assert.IsTrue(result.ToList().Count > 0);

        //}

        [TestMethod]
        [TestCategory("Process")]
        public void Process_GetEntryPoint_AlwaysReturnsValidWebForm()
        {

            _mockRepository = new SqlProcessRepository();

            var process = _mockRepository.GetProcess(1);
            var firstStep = process.GetEntryPoint();
            Assert.IsNotNull(firstStep);
            Assert.IsInstanceOfType(firstStep, typeof(ProcessFormElement));

        }
    }
}
