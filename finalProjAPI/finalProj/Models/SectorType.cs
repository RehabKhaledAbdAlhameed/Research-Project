using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace finalProj.Models
{
    public class SectorType
    {
        public int Sec_id { get; set; }
     
        public string Sec_Name { get; set; }

        public string IMG_URL { get; set; }

        public string Sec_Desc { get; set; }

        public virtual ICollection<Company> Companies { get; set; }
    }
}