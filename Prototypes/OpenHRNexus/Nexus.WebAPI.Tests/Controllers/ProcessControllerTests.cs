﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Classes;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Service.Services;
using Nexus.Sql_Repository;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using Nexus.WebAPI.Controllers;
using System;
using System.Collections.Generic;
using System.Security.Claims;

namespace Nexus.WebAPI.Tests.Controllers
{
    [TestClass]
    public class ProcessControllerTests
    {

        static IProcessRepository _mockRepository = new SqlProcessRepository();
        DataService _mockService = new DataService(_mockRepository);
        ProcessController _mockController;
        ClaimsIdentity _claims;

        [TestInitialize]
        public void TestInitialize()
        {
            _claims = new ClaimsIdentity();
            _claims.AddClaim(new Claim(ClaimTypes.Name, "testUser"));
            _claims.AddClaim(new Claim(ClaimTypes.NameIdentifier, "088C6A78-E14A-41B0-AD93-4FB7D3ADE96C"));

            _mockController = new ProcessController(_mockService, _claims, "fr-fr");

        }

        [TestMethod]
        public void GetPendingProcesses_ReturnsValidList()
        {
            var result = _mockController.GetPendingProcesses();
            Assert.IsInstanceOfType(result, typeof(IEnumerable<ProcessInFlow>));

        }


        [TestMethod]
        public void GetProcessStep_StartsNewProcessIfNoStepIdGiven()
        {
            var result = (List<WebFormModel>)_mockController.GetProcessStep(2, null);
            Assert.IsTrue(result.Count == 1);
        }

        [TestMethod]
        public void GetProcessStep_ConstructorReturnsNotNull()
        {
            var result = _mockController.GetProcessStep(1, null);
            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void GetProcessStep_HandlesInvalidUser()
        {
            _claims = new ClaimsIdentity();
            _claims.AddClaim(new Claim(ClaimTypes.Name, "NoSuchUser"));
            _claims.AddClaim(new Claim(ClaimTypes.NameIdentifier, "00000000-0000-0000-0000-000000000000"));

            _mockController = new ProcessController(_mockService, _claims, "en-GB");

            var result = (List<WebFormModel>)_mockController.GetProcessStep(2, null);

            Assert.IsTrue(result is IEnumerable<WebFormModel>);
            Assert.IsTrue(result.Count == 0);

        }

        [TestMethod]
        public void InstantiateProcess_ContainsWebformModels()
        {
            var getID = 1;
            var result = (List<WebFormModel>)_mockController.GetProcessStep(getID, null);

            Assert.IsTrue(result is IEnumerable<WebFormModel>);
            Assert.IsTrue(result[0].id == getID);

        }

        [TestMethod]
        public void PostProcessStep_ResponseReceived()
        {

            var field = new WebFormField { sequence = 1, columnid = 2, value = "Smith" };

            var form = new WebFormModel
            {
                stepid = Guid.NewGuid(),
                fields = new List<WebFormField>() {
                    new WebFormField { id=1, sequence = 1, columnid = 1, value = "John" },
                    new WebFormField { id=1, sequence = 2, columnid = 2, value = "Smith" },
                }
            };


            var result = _mockController.PostProcessStep(form);
            Assert.IsTrue(result is ProcessStepResponse);

        }




    }
}