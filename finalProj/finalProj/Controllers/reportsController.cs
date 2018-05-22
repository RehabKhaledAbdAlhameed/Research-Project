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
        public IEnumerable<Report> Reports (string Tag_Comp)
        {
            using (DBDataModel db = new DBDataModel())
            {
                return db.Reports.Where(e=>e.Tag_Comp == Tag_Comp).ToList();
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
