using finalProj.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace finalProj.Controllers
{
    public class NonPeriodicController : ApiController
    {
        [HttpGet]
        public IEnumerable<Periodic_Date_Live> get ()
        {
            using (DBDataModel db = new DBDataModel() )
            {
                return db.Periodic_Date_Live.ToList();
            }
        }
        // [HttpPost]
        public Periodic_Date_Live get( int ID)
        {
            using (DBDataModel db = new DBDataModel())
            {
                return db.Periodic_Date_Live.FirstOrDefault(c => c.ID == ID);
            }
        }
    }
}
