using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Classes;
using Nexus.Common.Interfaces;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Service.Services;
using System;
using System.Collections.Generic;
using System.Net.Mail;

namespace Nexus.Service.Tests
{
    [TestClass]
    public class SubmitStepTests
    {

        static IProcessRepository _mockRepository; // = new SqlProcessRepository();
        static ITranslation _mockDictionary; // = new SqlDictionaryRepository();
        DataService _mockService = new DataService(_mockRepository, _mockDictionary);

        private static ProcessEmailTemplate _mockEmailTemplate = new ProcessEmailTemplate()
        {
            Body = "<html>{button3}</html>",
            Subject = "Test Subject",
            Destinations = new EmailAddressCollection()
            {
                To = "lofty@asr.co.uk",
                From = "lofty@asr.co.uk"
            },
            FollowOnActions = new List<WebFormButtonModel>()
            {
                new WebFormButtonModel() {id = 3, TargetStep = new Guid("00000000-0000-0000-0000-000000000001"), TargetUrl = "helloduckyworld.com" },
                new WebFormButtonModel() {id = 4, TargetStep = new Guid("AAAAAAAA-AAAA-AAAA-AAAA-000000000002"), TargetUrl = "helloduckyworld.com" },
                new WebFormButtonModel() {id = 5, TargetStep = new Guid("AAAAAAAA-AAAA-AAAA-AAAA-000000000003"), TargetUrl = "helloduckyworld.com" },
                new WebFormButtonModel() {id = 6, TargetStep = new Guid("AAAAAAAA-AAAA-AAAA-AAAA-000000000004"), TargetUrl = "helloduckyworld.com" }
            }
        };


        [TestMethod]
        [TestCategory("Submit Step")]
        public void DataService_FollowOnActionsApplyAuthentication()
        {
            Assert.Fail("The authentication service needs implementing as an interface.");
        }

        [TestMethod]
        [TestCategory("Submit Step")]
        public void DataService_MailMessageTypeIsGenerated()
        {

            var variables = new Dictionary<string, object>()
                {
                    {"we_18_9", DateTime.Now },
                    {"we_23_10", "AM" },
                    {"we_22_8", 2 }
                };

            _mockEmailTemplate.Variables = variables;

            var result = _mockEmailTemplate.GenerateMailMessage();
            Assert.IsInstanceOfType(result, typeof(MailMessage));

        }


    }
}
