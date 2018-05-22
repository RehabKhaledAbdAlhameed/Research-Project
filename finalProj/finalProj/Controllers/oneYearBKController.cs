using finalProj.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace finalProj.Controllers
{
    public class oneYearBKController : ApiController
    {
        //[HttpPost]
        public IEnumerable<Historical_Price> get(int Comp_id)
        {
            using (DBDataModel db= new DBDataModel())
            {
                return db.Historical_Price.Where(e => e.Comp_id == Comp_id).ToList();
            }
        }
    }
}
