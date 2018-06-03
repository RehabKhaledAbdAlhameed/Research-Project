using finalProj.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web;
using System.Web.Mvc;
using System.Data.Entity;
using System.Data;

namespace finalProj.Controllers
{
    public class IndustriesController : ApiController
    {
        // GET api/<controller>
        public List<IndustryType> Get()
        {
            using (DBDataModel db = new DBDataModel())
            {
                //var a = db.Industries.Include(c => c.Companies).ToList();

                //return a;
                var result = (from i in db.Industries
                              join s in db.Sectors on i.Sec_id equals s.Sec_id
                              select new IndustryType
                              {
                                  Ind_id = i.Ind_id,
                                  Ind_Name = i.Ind_Name,
                                  Sec_Name = s.Sec_Name,
                                  IMG_URL = i.IMG_URL

                              }).ToList();
                return result;
            }
        }
        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("IndustryCompanies")]
        public List<Industry> GetCompanies()
        {
            using (DBDataModel db = new DBDataModel())
            {
                var a = db.Industries.Include(c => c.Companies).ToList();

                return a;
            }
        }
    }
}
