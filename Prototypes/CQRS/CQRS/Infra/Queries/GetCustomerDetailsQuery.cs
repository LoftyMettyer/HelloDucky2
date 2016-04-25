using Core;
using Core.Queries;

namespace Infra.Queries
{
	public class GetCustomerDetailsQuery : IQuery<Customer>
	{
		public int id { get; set; }
	}
}