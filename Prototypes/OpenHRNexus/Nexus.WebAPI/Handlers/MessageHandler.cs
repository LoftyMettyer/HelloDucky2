using System;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Nexus.WebAPI.Handlers {
	public abstract class MessageHandler : DelegatingHandler {
		protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,
			CancellationToken cancellationToken) {
			//Generate a correlation id that will help us pair the request with the response
			var correlationId = String.Format("[{0}] - {1}{2}", DateTime.Now, DateTime.Now.Ticks, Thread.CurrentThread.ManagedThreadId);

			//Get the request so we can log it later (if configured to do so)
			var requestInfo = String.Format("{0} {1}", request.Method, request.RequestUri);
			var requestMessage = await request.Content.ReadAsByteArrayAsync();
			await IncommingMessageAsync(correlationId, requestInfo, requestMessage);

			//Call the base method so we don't break the request/response flow
			var response = await base.SendAsync(request, cancellationToken);

			//Get the response so we can log it later (if configured to do so)
			byte[] responseMessage;
			if (response.IsSuccessStatusCode)
				responseMessage = await response.Content.ReadAsByteArrayAsync();
			else
				responseMessage = Encoding.UTF8.GetBytes(response.ReasonPhrase);

			await OutgoingMessageAsync(correlationId, requestInfo, responseMessage);

			//
			return response;
		}

		protected abstract Task IncommingMessageAsync(string correlationId, string requestInfo, byte[] message);
		protected abstract Task OutgoingMessageAsync(string correlationId, string requestInfo, byte[] message);
	}
}
