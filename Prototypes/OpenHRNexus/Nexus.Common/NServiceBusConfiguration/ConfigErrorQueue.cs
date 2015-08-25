﻿using NServiceBus.Config;
using NServiceBus.Config.ConfigurationSource;

namespace Nexus.Common.NServiceBusConfiguration {
	class ConfigErrorQueue : IProvideConfiguration<MessageForwardingInCaseOfFaultConfig> {
		public MessageForwardingInCaseOfFaultConfig GetConfiguration() {
			return new MessageForwardingInCaseOfFaultConfig {
				ErrorQueue = "error"
			};
		}
	}
}
