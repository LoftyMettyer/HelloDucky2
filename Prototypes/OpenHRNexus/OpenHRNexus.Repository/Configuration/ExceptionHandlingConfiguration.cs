﻿using OpenHRNexus.Repository.Globals;

namespace OpenHRNexus.Repository.Configuration {
	public class ExceptionHandlingConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new OpenHRNexus.EnterpriseServices.ExceptionHandling.NexusExceptionManager("OpenHRNexus.Repository", true, true);
		}
	}
}
