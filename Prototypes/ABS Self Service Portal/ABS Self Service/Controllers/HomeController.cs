using ABS_Self_Service.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;


namespace ABS_Self_Service.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {

            var currentUserId = User.Identity.GetUserId();
            if (currentUserId != null) {
                var manager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));
                var currentUser = manager.FindById(currentUserId);
                ViewBag.OpenHRID = currentUser.OpenHRID;

                //get some widgets
                widgetsModel m = new widgetsModel();
                var model = m.LoadModel(Server.MapPath("~/Config/PortalConfig.xml"));
                return View(model);
            }

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}