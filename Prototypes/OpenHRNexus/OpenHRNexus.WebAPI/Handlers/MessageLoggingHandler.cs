using System;
using System.Text;
using System.Threading.Tasks;
using OpenHRNexus.WebAPI.Globals;

namespace OpenHRNexus.WebAPI.Handlers {
	public class MessageLoggingHandler : MessageHandler {
		protected override async Task IncommingMessageAsync(string correlationId, string requestInfo, byte[] message) {
			await Task.Run(() =>
				Global.NexusLoggingManager.LogWriter.Write(
					String.Format("{0} - Request: {1}\r\n{2}", correlationId, requestInfo, Encoding.UTF8.GetString(message))
				)
			);
		}

		protected override async Task OutgoingMessageAsync(string correlationId, string requestInfo, byte[] message) {
			await Task.Run(() =>
				Global.NexusLoggingManager.LogWriter.Write(
					String.Format("{0} - Response: {1}\r\n{2}", correlationId, requestInfo, Encoding.UTF8.GetString(message))
				)
			);
		}
	}
}
