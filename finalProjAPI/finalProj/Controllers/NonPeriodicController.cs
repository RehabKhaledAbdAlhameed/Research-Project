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
        public IEnumerable<NonPeriodic_Data_Live> get ()
        {
            using (DBDataModel db = new DBDataModel() )
            {
                return db.NonPeriodic_Data_Live.ToList();
            }
        }
        // [HttpPost]
        public NonPeriodic_Data_Live  get( int ID)
        {
            using (DBDataModel db = new DBDataModel())
            {
                return db.NonPeriodic_Data_Live.FirstOrDefault(c => c.comp_id == ID);
            }
        }
    }
}
