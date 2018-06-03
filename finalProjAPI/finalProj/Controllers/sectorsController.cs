using finalProj.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web;
using System.Data.Entity;
using System.Data;

namespace finalProj.Controllers
{
    public class SectorsController : ApiController
    {
        [HttpGet]
        public List<SectorType> GetSector()
        {
            using (DBDataModel db = new DBDataModel())
            {
                var sectors = (from s in db.Sectors
                               select new SectorType
                               {
                                   Sec_id=s.Sec_id,
                                   Sec_Name = s.Sec_Name,
                                   Sec_Desc=s.Sec_Desc,
                                   IMG_URL=s.IMG_URL,
                                   //Industries=s.Industries
                                   //Companies=s.Companies
                               }).ToList();

                return sectors;
            }
        }


        // GET api/<controller>
        [HttpGet]
        [Route("SectorCompanies")]
        public List<Sector> GetCompanies()
        {
            using (DBDataModel db = new DBDataModel())
            {
                var a = db.Sectors.Include(c => c.Companies).ToList();

                return a;
            }
        }
    }
}
