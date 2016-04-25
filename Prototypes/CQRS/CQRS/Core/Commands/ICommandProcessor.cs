namespace Core.Commands
{
	public interface ICommandProcessor
	{
		TResult Process<TResult>(ICustomCommand command);
	}
}