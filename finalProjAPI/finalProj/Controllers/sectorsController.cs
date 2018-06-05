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
using System.Web.Http.Cors;

namespace finalProj.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
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

        //http://researchwebapi.azurewebsites.net/SectorIndustries?id=6
        [HttpGet]
        [Route("SectorIndustries")]
        public List<IndustryType> GetIndustries(int id)
        {
            using (DBDataModel db = new DBDataModel())
            {
                var result = (from i in db.Industries
                              join s in db.Sectors on i.Sec_id equals s.Sec_id
                              where i.Sec_id== id
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


        [HttpGet]
        [Route("Industry_Specifec_Companies")]

        public List<CompanyType> GetIndustryCompanies(int id)
        {
            using (DBDataModel db = new DBDataModel())
            {
                var companiesList = (from c in db.Companies
                                     join s in db.Sectors on c.Sec_id equals s.Sec_id
                                     join i in db.Industries on c.Ind_id equals i.Ind_id
                                     where c.Ind_id==id
                                     select new CompanyType
                                     {
                                         Comp_id = c.Comp_id,
                                         Comp_Name = c.Comp_Name,
                                         Reu_Code = c.Reu_Code,
                                         Com_Country = c.Com_Country,
                                         Trd_Curr = c.Trd_Curr,
                                         Curr_Out_Shares = (decimal)c.Curr_Out_Shares,
                                         IMG_URL = c.IMG_URL,
                                         Comp_Country = c.Com_Country,
                                         Sec_Name = c.Sector.Sec_Name,
                                         Ind_Name = c.Industry.Ind_Name
                                     }).ToList();

                return companiesList;


            }
        }

    }
}
/*
 

        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("Industry_Specifec_Companies")]
        public List<Company> GetIndustryCompanies(int id)
        {
            using (DBDataModel db = new DBDataModel())
            {
                var a = (from x in db.Companies
                         where x.Ind_id == id
                         select x).ToList();

                return a;
            }
        }

*/