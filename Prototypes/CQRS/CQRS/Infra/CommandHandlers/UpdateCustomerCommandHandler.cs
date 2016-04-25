using System.Data.Entity.Migrations;
using Core;
using Core.Commands;
using Infra.Commands;

namespace Infra.CommandHandlers
{
	public class UpdateCustomerCommandHandler : ICommandHandler<UpdateCustomerCommand, Customer>
	{
		private readonly ApplicationContext db;

		public UpdateCustomerCommandHandler()
		{
			db = new ApplicationContext();
		}

		public Customer Handle(UpdateCustomerCommand command)
		{
			Customer original = db.Customers.Find(command.Customer.Id);

			if (original != null)
			{
				original.Name = command.Customer.Name;
				original.Address = command.Customer.Address;
				db.Customers.AddOrUpdate(original);
				db.SaveChanges();
			}
			return command.Customer;

			//return null;
		}
	}
}