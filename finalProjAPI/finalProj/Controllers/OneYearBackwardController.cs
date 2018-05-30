using finalProj.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;



namespace finalProj.Controllers
{
    public class OneYearBackwardController : ApiController
    {
        // GET api/<controller>
       

        // GET api/<controller>/5
        
        public List<CompPrice> Get(int compid1)
        {
            using (DBDataModel db = new DBDataModel())
            {
                var x = DateTime.Now.AddMonths(-12);
                var q = (from c in db.Companies
                         join H in db.Historical_Price on c.Comp_id equals H.Comp_id
                         where H.Date >= x && c.Comp_id == compid1
                         select new CompPrice { Comp_Name = c.Comp_Name, Price_val = H.Price_val, }).ToList();
                return q;
            }
        }
      
    }
}