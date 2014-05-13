using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.ModelBinding;
using System.Web.Http.OData;
using System.Web.Http.OData.Routing;
using WebAPI3.Models;

namespace WebAPI3.Controllers
{
    /*
    To add a route for this controller, merge these statements into the Register method of the WebApiConfig class. Note that OData URLs are case sensitive.

    using System.Web.Http.OData.Builder;
    using WebAPI3.Models;
    ODataConventionModelBuilder builder = new ODataConventionModelBuilder();
    builder.EntitySet<Salary>("Salary");
    config.Routes.MapODataRoute("odata", "odata", builder.GetEdmModel());
    */
    public class SalaryController : ODataController
    {
        private SalaryContext db = new SalaryContext();

        // GET odata/Salary
        [Queryable]
        public IQueryable<Salary> GetSalary()
        {
            return db.Salaries;
        }

        // GET odata/Salary(5)
        [Queryable]
        public SingleResult<Salary> GetSalary([FromODataUri] int key)
        {
            return SingleResult.Create(db.Salaries.Where(salary => salary.ID == key));
        }

        // PUT odata/Salary(5)
        public IHttpActionResult Put([FromODataUri] int key, Salary salary)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (key != salary.ID)
            {
                return BadRequest();
            }

            db.Entry(salary).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!SalaryExists(key))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return Updated(salary);
        }

        // POST odata/Salary
        public IHttpActionResult Post(Salary salary)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.Salaries.Add(salary);
            db.SaveChanges();

            return Created(salary);
        }

        // PATCH odata/Salary(5)
        [AcceptVerbs("PATCH", "MERGE")]
        public IHttpActionResult Patch([FromODataUri] int key, Delta<Salary> patch)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            Salary salary = db.Salaries.Find(key);
            if (salary == null)
            {
                return NotFound();
            }

            patch.Patch(salary);

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!SalaryExists(key))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return Updated(salary);
        }

        // DELETE odata/Salary(5)
        public IHttpActionResult Delete([FromODataUri] int key)
        {
            Salary salary = db.Salaries.Find(key);
            if (salary == null)
            {
                return NotFound();
            }

            db.Salaries.Remove(salary);
            db.SaveChanges();

            return StatusCode(HttpStatusCode.NoContent);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool SalaryExists(int key)
        {
            return db.Salaries.Count(e => e.ID == key) > 0;
        }
    }
}
