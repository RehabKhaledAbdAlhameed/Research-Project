using finalProj.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace finalProj.Controllers
{
    public class reportsController : ApiController
    {
        // [HttpPost]
        public Report Reports (int id)
        {
            using (DBDataModel db = new DBDataModel())
            {
                return db.Reports.Where(e=>e.comp_id == id).FirstOrDefault();
            }
        }
        [HttpGet]
        public IEnumerable<Report> GetReports()
        {
            using (DBDataModel db = new DBDataModel())
            {
                return db.Reports.ToList();
            }
        }

    }
}
