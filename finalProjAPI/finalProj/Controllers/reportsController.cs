using finalProj.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;

namespace finalProj.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class reportsController : ApiController
    {
        [HttpGet]
        public IEnumerable<Report> Reports (int id)
        {
            using (DBDataModel db = new DBDataModel())
            {
                return db.Reports.Where(p => p.comp_id == id).ToList();
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
