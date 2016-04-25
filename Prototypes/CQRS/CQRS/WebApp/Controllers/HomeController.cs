using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Core;
using Infra.Commands;
using Infra.Processor;
using Infra.Queries;

namespace WebApp.Controllers
{
	public class HomeController : Controller
	{
		// GET:Home/Index
		public ActionResult Index()
		{
			var query = new GetAllCustomerQuery();
			var queryProcessor = new QueryProcessor();
			IEnumerable<Customer> data = queryProcessor.Process(query);
			return View(data.ToList());
		}


		// GET:Home/Create
		public ActionResult Create()
		{
			return View();
		}

		// POST: Home/Create
		// To protect from overposting attacks, please enable the specific properties you want to bind to, for 
		// more details see http://go.microsoft.com/fwlink/?LinkId=317598.
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(Customer customer)
		{
			var commandProcessor = new CommandProcessor();
			var command = new CreateCustomerCommand
			{
				Customer = customer
			};

			var result = commandProcessor.Process<Customer>(command);
			return RedirectToAction("Index");
		}

		// GET: Home/Create
		public ActionResult Edit(int Id = 0)
		{
			var query = new GetCustomerDetailsQuery {id = Id};
			var queryProcessor = new QueryProcessor();
			Customer customer = queryProcessor.Process(query);
			return View(customer);
		}

		// POST: Home/Create
		// To protect from overposting attacks, please enable the specific properties you want to bind to, for 
		// more details see http://go.microsoft.com/fwlink/?LinkId=317598.
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Edit(Customer customer)
		{
			var commandProcessor = new CommandProcessor();
			var command = new UpdateCustomerCommand
			{
				Customer = customer
			};

			var result = commandProcessor.Process<Customer>(command);
			return RedirectToAction("Index");
		}
	}
}