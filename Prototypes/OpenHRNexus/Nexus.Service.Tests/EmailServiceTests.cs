﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Classes;
using Nexus.Common.Enums;

namespace Nexus.Service.Tests {
	[TestClass]
	public class EmailServiceTests {
		[TestMethod]
		public void SendEmailSuccessfully() {
			//Commented this test out because sending email costs credits, since we are Postmark (https://postmarkapp.com), an email-sending free(ish) service

			//var emailService = new EmailService();
			//var result = emailService.Send("roberto.caballero@advancedcomputersoftware.com", "Test Subject", "Test Body");

			//Assert.AreEqual(result.Status, BusinessProcessStepStatus.EmailSuccessfullySent);
		}
	}
}