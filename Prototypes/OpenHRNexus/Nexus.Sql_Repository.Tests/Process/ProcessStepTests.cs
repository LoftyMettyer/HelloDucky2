using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Enums;
using Nexus.Common.Models;
using System;
using System.Collections.Generic;

namespace Nexus.Sql_Repository.Tests.Data
{
    [TestClass]
    public class ProcessStepTests
    {
        SqlProcessRepository _mockRepository;

        [TestInitialize]
        public void TestInitialize()
        {
            _mockRepository = new SqlProcessRepository();
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
            Assert.IsTrue(result.Status == ProcessStepStatus.Success);

        }

        [TestMethod]
        public void ProcessStep_CommitStepForAllDataTypes()
        {

            var form = new WebFormModel
            {
                stepid = Guid.NewGuid(),
                fields = new List<WebFormField>() {
                    new WebFormField { id=1, sequence = 1, columnid = 5, value = DateTime.Now.ToString()},
                    new WebFormField { id=1, sequence = 2, columnid = 13, value = 3.75.ToString() },
                    new WebFormField { id=1, sequence = 2, columnid = 4, value = "Jones" },
                    new WebFormField { id=1, sequence = 2, columnid = 7, value = true.ToString() },
                    new WebFormField { id=1, sequence = 2, columnid = 25, value = 0.ToString() }
                }
            };

            var userId = new Guid("088C6A78-E14A-41B0-AD93-4FB7D3ADE96C");

            var result = _mockRepository.CommitStep(form.stepid, userId, form);
            Assert.IsTrue(result.Status == ProcessStepStatus.Success);

        }

        [TestMethod]
        public void ProcessStep_RecordProcessStep_ReturnsGuid()
        {

            var form = new WebForm
            {
                Fields = new List<WebFormField>() {
                    new WebFormField { id=1, sequence = 1, columnid = 5, value = DateTime.Now.ToString()},
                    new WebFormField { id=1, sequence = 2, columnid = 13, value = 3.75.ToString() },
                    new WebFormField { id=1, sequence = 2, columnid = 4, value = "Jones" },
                    new WebFormField { id=1, sequence = 2, columnid = 7, value = true.ToString() },
                    new WebFormField { id=1, sequence = 2, columnid = 25, value = 0.ToString() }
                }
            };

            var userId = new Guid("088C6A78-E14A-41B0-AD93-4FB7D3ADE96C");

            var result = _mockRepository.RecordProcessStepForUser(form, userId);
            Assert.IsInstanceOfType(result, typeof(Guid));
            Assert.AreNotEqual(result, Guid.Empty);

        }

        [TestMethod]
        public void ProcessStep_RecordProcessStep_EmptyFormReturnsEmptyGuid()
        {
            var result = _mockRepository.RecordProcessStepForUser(null, Guid.Empty);
            Assert.AreEqual(result, Guid.Empty);
        }

    }
}
