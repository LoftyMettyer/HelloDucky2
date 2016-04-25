namespace Core.Queries
{
	public interface IQueryProcessor
	{
		TResult Process<TResult>(IQuery<TResult> query);
	}
}