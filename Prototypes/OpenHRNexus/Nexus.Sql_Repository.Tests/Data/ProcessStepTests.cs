﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Enums;
using Nexus.Common.Models;
using System;
using System.Collections.Generic;

namespace Nexus.Sql_Repository.Tests.Data
{
    [TestClass]
    public class ProcessStepTests
    {
        SqlDataRepository _mockRepository;

        [TestInitialize]
        public void TestInitialize()
        {
            _mockRepository = new SqlDataRepository();
        }

        [TestMethod]
        public void ProcessStep_CommitToTransactionLogReturnsSuccess()
        {


            var form = new WebFormModel
            {
                stepid = Guid.NewGuid(),
                fields = new List<WebFormField>() {
                    new WebFormField { id=1, sequence = 1, columnid = 1, value = "John" },
                    new WebFormField { id=1, sequence = 2, columnid = 2, value = "Smith" },
                }
            };

            var userId =  new Guid("088C6A78-E14A-41B0-AD93-4FB7D3ADE96C");

            var result = _mockRepository.CommitStep(form.stepid, userId, form);
            Assert.IsTrue(result.Status == BusinessProcessStepStatus.Success);

        }
    }
}