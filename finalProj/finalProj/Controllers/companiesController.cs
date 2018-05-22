using finalProj.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace finalProj.Controllers
{
    public class companiesController : ApiController
    {
        [HttpGet]
        public IEnumerable<Company> GetCompany()
        {
            using (DBDataModel db = new DBDataModel())
            {
                return db.Companies.ToList();
            }
        }
    }
}
