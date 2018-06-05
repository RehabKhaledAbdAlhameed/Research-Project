using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.Entity;
using meunsubmenu.Models;
using System.Data;
using System.Net;


namespace meunsubmenu.Controllers
{
    public class Sectorcompanymenu : Controller
    {
        private Model1 db1 = new Model1();
        public ActionResult Index()
        {

            using (var db = new Model1())
            {

                var d = db.Sectors.Include(c => c.Companies).ToList();
                return View(d);
            }
            //  return View();
        }
        public ActionResult Details(int? id)
        {
            using (var db = new Model1())
            {
                if (id == null)
                {
                    return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
                }
                Company company = db1.Companies.Find(id);

                if (company == null)
                {
                    return HttpNotFound();
                }
                return View(company);
            }
        }


    }
}