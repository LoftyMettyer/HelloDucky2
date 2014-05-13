using ProductsApp.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Web.Configuration;
using System.Web.Http;

namespace ProductsApp.Controllers
{
	public class ProductsController : ApiController
	{

      //  static string connStr = ConfigurationManager.ConnectionStrings["OpenHR"].ConnectionString;
     //   SqlConnection con = new SqlConnection(connStr);




		Product[] products = new Product[] 
        { 
            new Product { Id = 1, Name = "Tomato Soup", Category = "Groceries", Price = 1 }, 
            new Product { Id = 2, Name = "Yo-yo", Category = "Toys", Price = 3.75M }, 
            new Product { Id = 3, Name = "Hammer", Category = "Hardware", Price = 16.99M } 
        };

		public IEnumerable<Product> GetAllProducts()
		{



			return products;
		}

		public IHttpActionResult GetProduct(int id)
		{

            try
            {
                int i = 3;

            }
            catch
            {

            }


      //      SqlDataAdapter objAdaptor = new SqlDataAdapter();
        //    SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["OpenHR"].ToString);

            //        string sConn = WebConfigurationManager.ConnectionStrings["OpenHR"];


          //  connStr = ConfigurationManager.ConnectionStrings["OpenHR"];

			var product = products.FirstOrDefault((p) => p.Id == id);
			if (product == null)
			{
				return NotFound();
			}
			return Ok(product);
		}
	}
}