using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using finalProjMVC.Models;
using System.Net.Http;
using Newtonsoft.Json;

namespace finalProjMVC.Controllers
{
    public class CompaniesController : Controller
    {
        public HttpClient client ;
        static List<CompanyType> companies   = new List<CompanyType>();
        public CompaniesController()
        {
            client = new HttpClient();
            client.BaseAddress = new Uri("http://localhost:51683/api/");
        }

        // GET: Companies
        public ActionResult Index()
        {

            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            var res = client.GetAsync("companies").Result;
            if (res.IsSuccessStatusCode)
            {
                var msg = res.Content.ReadAsStringAsync().Result;
                JsonSerializerSettings jsonsettings = new JsonSerializerSettings();
                var CompaniesList = JsonConvert.DeserializeObject<List<CompanyType>>(msg, jsonsettings);
                companies = CompaniesList.ToList();
                return View(companies);
            }
            return View();
        }

        // GET: Companies/Details/5
        public ActionResult Details(int? id)
        {

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

            var res =client.GetAsync($"companies/{id}").Result;
            var rep= client.GetAsync($"reports/{id}").Result;
            if (res.IsSuccessStatusCode)
            {
                var msgr = rep.Content.ReadAsStringAsync().Result;
                JsonSerializerSettings jsonsettingsr = new JsonSerializerSettings();
                ViewBag.reports = JsonConvert.DeserializeObject<List<Report>>(msgr, jsonsettingsr).ToList();

                var msgc = res.Content.ReadAsStringAsync().Result;
                JsonSerializerSettings jsonsettingsc = new JsonSerializerSettings();
                var Company = JsonConvert.DeserializeObject<CompanyType>(msgc, jsonsettingsc);
                return View(Company);
            }
            return View();
        }
        
    }
}
