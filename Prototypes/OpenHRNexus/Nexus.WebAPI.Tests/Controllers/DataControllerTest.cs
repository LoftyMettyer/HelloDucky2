using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Service.Services;
using Nexus.WebAPI.Controllers;
using Nexus.Common.Models;
using Nexus.Sql_Repository;
using System.Security.Claims;
using System;
using Nexus.Common.Classes;
using System.Web.Mvc;
using Nexus.Common.Enums;

namespace Nexus.WebAPI.Tests.Controllers
{
	[TestClass]
	public class DataControllerTest
	{

        SqlProcessRepository _mockRepository;
        DataService _mockService;
        DataController _mockController;
        ClaimsIdentity _claims;

        [TestInitialize]
		public void TestInitialize()
		{
            _claims = new ClaimsIdentity();
            _claims.AddClaim(new Claim(ClaimTypes.Name, "testUser"));
            _claims.AddClaim(new Claim(ClaimTypes.NameIdentifier, "088C6A78-E14A-41B0-AD93-4FB7D3ADE96C"));

            _mockRepository = new SqlProcessRepository();
            _mockService = new DataService(_mockRepository);
            _mockController = new DataController(_mockService, _claims, "fr-fr");

        }

        [TestMethod]
        [Description("Building up a sql statement can cause errors if the column is included in the select multiple times. Ensure that we handle this.")]
        public void InstantiateProcess_HandlesTheSameColumnMultipleTimes()
        {
            Assert.Fail("Not yet implemented");
        }




        //[TestMethod]
        //public void SubmitStep_EmailSendsSuccessfully()
        //{

        //    var form = new WebFormModel
        //    {
        //        stepid = Guid.NewGuid(),
        //        fields = new List<WebFormField>() {
        //            new WebFormField { id=1, sequence = 1, columnid = 1, value = "John" },
        //            new WebFormField { id=1, sequence = 2, columnid = 2, value = "Smith" },
        //        }
        //    };

        //    var result = _mockController.SubmitStep(form);
        //    Assert.IsTrue(result is ProcessStepResponse);
        //    Assert.IsTrue(result.Status == ProcessStepStatus.EmailSuccessfullySent);

        //}

        [TestMethod]
        public void GetCalendarData_GetsDataForTodaysRange()
        {

            var result = (List<CalendarEventModel>)_mockController.GetCalendarData("absence", DateTime.Now, DateTime.Now);
            Assert.IsTrue(result is IEnumerable<CalendarEventModel>);

        }



    }


}
