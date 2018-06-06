
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
    public class SectorsController : Controller
    {
        public HttpClient client;
        public SectorsController()
        {
            client = new HttpClient();
            client.BaseAddress = new Uri("http://researchwebapi.azurewebsites.net/");
        }
        // GET: Sectors
        public ActionResult Index()
        {
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            var res = client.GetAsync("SectorCompanies").Result;
            if(res.IsSuccessStatusCode)
            {
                var msg = res.Content.ReadAsStringAsync().Result;
                JsonSerializerSettings jsonsetting = new JsonSerializerSettings();
                var SectorsList = JsonConvert.DeserializeObject<List<Sector>>(msg,jsonsetting).ToList();
                return View(SectorsList);
            }

            return View();
        }
    }
}