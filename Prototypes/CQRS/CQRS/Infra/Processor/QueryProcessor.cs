using System;
using Autofac;
using Core.Queries;
using Infra.IOCContainer;

namespace Infra.Processor
{
	public class QueryProcessor : IQueryProcessor
	{
		/// <summary>
		///     Automatically figures out which QueryHandler belongs to the given Query, instantiates it, and returns the result of
		///     running the Query on that QueryHandler.
		/// </summary>
		/// <typeparam name="TResult">The type of result to return. This can be infered from Query given.</typeparam>
		/// <param name="query">The query to return to use as basis for finding a suitable QueryHandler and returning the result.</param>
		/// <returns></returns>
		public TResult Process<TResult>(IQuery<TResult> query)
		{
			Type handlerType = typeof (IQueryHandler<,>)
				.MakeGenericType(query.GetType(), typeof (TResult));

			dynamic handler = IoCBuilder.Container.Resolve(handlerType);
			return handler.Handle((dynamic) query);
		}
	}
}