using Core;
using Core.Commands;

namespace Infra.Commands
{
	public class UpdateCustomerCommand : ICustomCommand
	{
		public Customer Customer { get; set; }
	}
}