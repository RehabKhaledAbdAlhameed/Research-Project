using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Http.Description;
using finalProj.Models;

namespace finalProj.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class CompaniesController : ApiController
    {

        [HttpGet]
        public List<CompanyType> GetCompany()
        {
            using (DBDataModel db = new DBDataModel())
            {
                var companiesList = (from c in db.Companies
                                     join s in db.Sectors on c.Sec_id equals s.Sec_id
                                     join i in db.Industries on c.Ind_id equals i.Ind_id
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

        [HttpGet]
       // [Route("compid")]
        public CompanyType GetCompany(int id)
        {
            using (DBDataModel db = new DBDataModel())
            {
                var company = (from c in db.Companies
                               join s in db.Sectors on c.Sec_id equals s.Sec_id
                               join i in db.Industries on c.Ind_id equals i.Ind_id
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
                               }).Where(p => p.Comp_id == id).FirstOrDefault();

                return company;
            }
        }
    }
}