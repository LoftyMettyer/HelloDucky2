using System;
using Autofac;
using Core.Commands;
using Infra.IOCContainer;

namespace Infra.Processor
{
	public class CommandProcessor : ICommandProcessor
	{
		public TResult Process<TResult>(ICustomCommand command)
		{
			Type handlerType = typeof (ICommandHandler<,>).MakeGenericType(command.GetType(), typeof (TResult));
			dynamic handler = IoCBuilder.Container.Resolve(handlerType);
			return handler.Handle((dynamic) command);
		}
	}
}