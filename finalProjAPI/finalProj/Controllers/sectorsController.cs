
using finalProj.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace finalProj.Controllers
{
    public class sectorsController : ApiController
    {
        [HttpGet]
        public List<SectorType> GetSector()
        {
            using (DBDataModel db = new DBDataModel())
            {
                var sectors = (from s in db.Sectors
                               select new SectorType
                               {
                                   Sec_Name = s.Sec_Name
                               }).ToList();
                return sectors.ToList();
            }
        }
    }
}
