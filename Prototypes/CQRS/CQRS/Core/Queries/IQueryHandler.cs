namespace Core.Queries
{
	/// <summary>
	///     Defines a test-friendly query handler that can be automatically instantiated and populated with dependency
	///     injection.
	/// </summary>
	/// <typeparam name="TQuery">The type of query that this query handler should handle. For example, GetEmployeeByIdQuery.</typeparam>
	/// <typeparam name="TResult">
	///     The type of result that the query returns. Must be the same as defined in the query itself.
	///     For example, Employee.
	/// </typeparam>
	public interface IQueryHandler<TQuery, TResult> where TQuery : IQuery<TResult>
	{
		TResult Handle(TQuery query);
	}
}