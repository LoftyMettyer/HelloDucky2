using Core;
using Core.Commands;
using Infra.Commands;

namespace Infra.CommandHandlers
{
	public class CreateCustomerCommandHandler : ICommandHandler<CreateCustomerCommand, Customer>
	{
		private readonly ApplicationContext db;

		public CreateCustomerCommandHandler()
		{
			db = new ApplicationContext();
		}

		public Customer Handle(CreateCustomerCommand command)
		{
			db.Customers.Add(command.Customer);
			db.SaveChanges();
			return command.Customer;
			//return null;
		}
	}
}