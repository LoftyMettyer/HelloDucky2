using System.Collections.Generic;
using System.Linq;
using Core;
using Core.Queries;
using Infra.Queries;

namespace Infra.QueryHandlers
{
	public class GetAllCustomerQueryHandler : IQueryHandler<GetAllCustomerQuery, IEnumerable<Customer>>
	{
		private readonly ApplicationContext context;

		public GetAllCustomerQueryHandler()
		{
			context = new ApplicationContext();
		}

		public IEnumerable<Customer> Handle(GetAllCustomerQuery query)
		{
			return context.Set<Customer>().ToList();
		}
	}
}