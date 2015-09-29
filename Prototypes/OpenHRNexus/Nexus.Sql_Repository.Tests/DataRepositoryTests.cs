using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Classes;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using System;
using System.Collections.Generic;

namespace Nexus.Sql_Repository.Tests
{
    [TestClass]
    public class DataRepositoryTests
    {
        private IProcessRepository _mockRepository = new SqlProcessRepository();

        [TestMethod]
        [TestCategory("Get Data")]
        public void Repository_GetDataDefinition_ReturnsType()
        {
            var result = _mockRepository.GetDataDefinition(1);
            Assert.IsInstanceOfType(result, typeof(Type), "Type class is not returned");
            Assert.IsNotNull(result.GetFields(), "No fields returned");
        }

        [TestMethod]
        [TestCategory("Submit Step")]
        public void Repository_RecordProcessStepForUser_IsSuccess()
        {
            var userId = Guid.Empty;
            var process = new Process();
            var submitForm = new WebFormDataModel
            {
                stepid = Guid.NewGuid(),
                data = new Dictionary<string, object>()
                {
                    {"we_18_9", DateTime.Now },
                    {"we_23_10", "AM" },
                    {"we_22_8", 2 }
                }
            };

            var result = _mockRepository.UpdateProcessWithUserVariables(process, submitForm, userId);
            Assert.IsInstanceOfType(result, typeof(WebFormDataModel));

        }


        [TestMethod]
        public void Repository_RecordProcessStep_VariablesAreSeriablizable()
        {


        }

    }
}
