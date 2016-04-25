namespace Core.Commands
{
	public interface ICommandHandler<TCommand, TResult> where TCommand : ICustomCommand
	{
		TResult Handle(TCommand command);
	}
}