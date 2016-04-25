using System.Linq;
using Core;
using Core.Queries;
using Infra.Queries;

namespace Infra.QueryHandlers
{
	public class GetCustomerDetailsQueryHandler : IQueryHandler<GetCustomerDetailsQuery, Customer>
	{
		private readonly ApplicationContext context;

		public GetCustomerDetailsQueryHandler()
		{
			context = new ApplicationContext();
		}

		public Customer Handle(GetCustomerDetailsQuery query)
		{
			return context.Set<Customer>().ToList().Find(x => x.Id == query.id);
		}
	}
}