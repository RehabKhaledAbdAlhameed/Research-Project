using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace finalProj.Models
{
    public class IndustryType
    {
        public int Ind_id { get; set; }

   
        public string Ind_Name { get; set; }

        public string Sec_Name { get; set; }


        public string IMG_URL { get; set; }

        public virtual ICollection<Company> Companies { get; set; }
    }
}