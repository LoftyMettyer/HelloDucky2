using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Classes;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Service.Services;
using Nexus.Sql_Repository;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using Nexus.WebAPI.Controllers;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace Nexus.WebAPI.Tests.Controllers
{
    [TestClass]
    public class ProcessControllerTests
    {

        static IProcessRepository _mockRepository = new SqlProcessRepository();
        static SqlDictionaryRepository _mockDictionary = new SqlDictionaryRepository();
        DataService _mockService = new DataService(_mockRepository, _mockDictionary);
        ProcessController _mockController;
        ClaimsIdentity _claims;

        [TestInitialize]
        public void TestInitialize()
        {
            _mockService.CallingURL = "http://nexus-advanced.azurewebsites.net/";
            _mockService.AuthenticationServiceURL = "http://abs16091/authenticationservice/";
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
            var result = (List<WebFormModel>)_mockController.GetProcessStep(2);
            Assert.IsTrue(result.Count == 1);
        }

        [TestMethod]
        public void GetProcessStep_ConstructorReturnsNotNull()
        {
            var result = _mockController.GetProcessStep(2);
            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void GetProcessStep_FormFieldsAreInCorrectSequence()
        {
            var result = (List<WebFormModel>)_mockController.GetProcessStep(2);
            Assert.IsNotNull(result);

            int sequence = 0;
            foreach (var field in result[0].fields)
            {
                Assert.IsTrue(field.sequence > sequence, "Form fields are not in correct sequence");
                sequence = field.sequence;
            }
        }


        [TestMethod]
        public void GetProcessStep_HandlesInvalidUser()
        {
            _claims = new ClaimsIdentity();
            _claims.AddClaim(new Claim(ClaimTypes.Name, "NoSuchUser"));
            _claims.AddClaim(new Claim(ClaimTypes.NameIdentifier, "00000000-0000-0000-0000-000000000000"));

            _mockController = new ProcessController(_mockService, _claims, "en-GB");

            var result = (List<WebFormModel>)_mockController.GetProcessStep(2);

            Assert.IsTrue(result is IEnumerable<WebFormModel>);
            Assert.IsTrue(result.Count == 0);

        }

        [TestMethod]
        public void GetProcessStep_ContainsWebformModels()
        {
            var getID = 1;
            var result = (List<WebFormModel>)_mockController.GetProcessStep(getID);

            Assert.IsTrue(result is IEnumerable<WebFormModel>);
            Assert.IsTrue(result[0].id == getID);

        }


        [TestMethod]
        public void PostProcessStep_ResponseReceived()
        {
            async_PostProcessStep_ResponseReceived();
        }


        [TestMethod]
        public async void async_PostProcessStep_ResponseReceived()
        {

            var form = new WebFormDataModel
            {
                stepid = Guid.NewGuid(),
                data = new Dictionary<string, object>()
                {
                    {"we_1_1", "John" },
                    { "we_2_2", "Smith" }
                }
            };

            var result = await _mockController.PostProcessStep(form);
            Assert.IsTrue(result is ProcessStepResponse);

        }


        [TestMethod]
        public void PostProcessStep_JsonIsSerialized()
        {

            var originalform = new WebFormDataModel
            {
                stepid = Guid.NewGuid(),
                data = new Dictionary<string, object>()
                {
                    {"we_1_1", "John" },
                    { "we_2_2", "Smith" }
                }
            };

            var serializedForm = new JavaScriptSerializer().Serialize(originalform);
            Assert.IsInstanceOfType(serializedForm, typeof(string));

            var reSerializedform = new JavaScriptSerializer().Deserialize<WebFormDataModel>(serializedForm);
            Assert.IsInstanceOfType(reSerializedform, typeof(WebFormDataModel));

            var result =  _mockController.PostProcessStep(reSerializedform);
            Assert.IsTrue(result is ProcessStepResponse, "Serialized object cannot be processed");

        }


        [TestMethod]
        public void PostProcessStep_EmailIsTriggered()
        {

            var form = new WebFormDataModel
            {
                stepid = Guid.NewGuid(),
                data = new Dictionary<string, object>()
                {
                    { "we_18_9", DateTime.Now },
                    { "we_19_11", DateTime.Now.AddDays(3) },
                    { "we_20_13", 2 },
                    { "we_21_14", "some notes" },
                    { "we_22_8", "3" },
                    { "we_23_10", "AM" },
                    { "we_24_12", "PM" },
                    { "we_25_15", false }

                }
            };

            var result = _mockController.PostProcessStep(form);
            Assert.IsTrue(result.Status == (TaskStatus) ProcessStepStatus.EmailSuccessfullySent);

        }

    }
  }
