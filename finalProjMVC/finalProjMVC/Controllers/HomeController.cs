using finalProjMVC.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace finalProjMVC.Controllers
{
    public class HomeController : Controller
    {

        DataModel db = new DataModel();
        public ActionResult Index()
        {
            ViewBag.comp = db.Companies.Count();
            ViewBag.ind = db.Industries.Count();
            ViewBag.sec = db.Sectors.Count();
            return View();
        }
    }
}