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
    public class IndCompController : ApiController
    {
        // GET api/<controller>
        public List<Industry> Get()
        {
            using (DBDataModel db = new DBDataModel())
            {
                var a = db.Industries.Include(c => c.Companies).ToList();

                return a;
            }
        }
    }
}