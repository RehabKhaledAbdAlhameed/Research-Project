using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.Entity;
using System.Data;
using System.Net;
using finalProjMVC.Models;
using System.Net.Http;
using Newtonsoft.Json;

namespace meunsubmenu.Controllers
{
    public class IndustriesController : Controller
    {
        public HttpClient client;
        public IndustriesController()
        {
            client = new HttpClient();
            client.BaseAddress= new Uri("http://localhost:51683/");
        }
        public ActionResult Index()
        {
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            var res = client.GetAsync("IndustryCompanies").Result;
            if(res.IsSuccessStatusCode)
            {
                var msg = res.Content.ReadAsStringAsync().Result;
                JsonSerializerSettings jsonsetting = new JsonSerializerSettings();
                var IndustriesList = JsonConvert.DeserializeObject<List<Industry>>(msg, jsonsetting).ToList();
                return View(IndustriesList);
            }
            return View();
        }
        //public ActionResult Details(int? id)
        //{
        //    using (var db = new DataModel())
        //    {
        //        if (id == null)
        //        {
        //            return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
        //        }
        //        Company company = db1.Companies.Find(id);
               
        //        if (company == null)
        //        {
        //            return HttpNotFound();
        //        }
        //        return View(company);
        //    }
        //}
     
    }
}