using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class pricesController : ApiController
    {
        Model1 model = new Model1();
        public List< CompPrice> getprices (string compname)
        {
            var q = (from c in model.Companies
                     join H in model.Historical_Price on c.Comp_id equals H.Comp_id
                     where c.Comp_Name.Contains(compname)
                     where H.Date.Year== (DateTime.Now.Year- 1)
                     select new CompPrice
                     {
                         Comp_Name = c.Comp_Name,
                         Price_val = H.Price_val,
                     }).ToList();

            return q;
        }
        
    }
}
