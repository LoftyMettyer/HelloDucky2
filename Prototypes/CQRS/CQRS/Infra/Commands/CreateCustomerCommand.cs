using Core;
using Core.Commands;

namespace Infra.Commands
{
	public class CreateCustomerCommand : ICustomCommand
	{
		public Customer Customer { get; set; }
	}
}