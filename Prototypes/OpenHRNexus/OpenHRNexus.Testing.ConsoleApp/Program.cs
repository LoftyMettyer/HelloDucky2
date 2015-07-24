using System;
using NServiceBus;
using OpenHRNexus.Common.Messaging.Commands;
using OpenHRNexus.Interfaces.Common;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Testing.ConsoleApp {
	internal class Program {
		private static void Main(string[] args) {
			// BusService_Test();
		}

		private static void BusService_Test() {
			var busConfiguration = new BusConfiguration();
			busConfiguration.EndpointName("OpenHRNexus.Testing.ConsoleApp");
			busConfiguration.UseSerialization<JsonSerializer>();
			busConfiguration.EnableInstallers();
			busConfiguration.UsePersistence<InMemoryPersistence>();

			using (IStartableBus startableBus = Bus.Create(busConfiguration)) {
				var bus = startableBus.Start();

				Console.Write("Press 'Enter' to send the first message");
				Console.ReadLine();
				var loginAttempt = new LoginAttemptCommand() { UserName = "peter", Password = "pan" }; //This user will succeed
				bus.Send("OpenHRNexus.Messaging.Publisher", loginAttempt);
				Console.WriteLine("Sending LoginAttemptCommand Command (successful one)");

				Console.Write("Press 'Enter' to send the second message");
				Console.ReadLine();
				loginAttempt = new LoginAttemptCommand() { UserName = "peter", Password = "pannetone" }; //This user will FAIL
				bus.Send("OpenHRNexus.Messaging.Publisher", loginAttempt);
				Console.WriteLine("Sending LoginAttemptCommand Command (unsuccessful one)");

				Console.ReadLine();
			}
		}
	}
}