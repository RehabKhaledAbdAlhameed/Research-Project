using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace finalProj.Models
{
    public class CompanyType
    {
        public int Comp_id { get; set; }
        public string Comp_Name { get; set; }
        public string Reu_Code { get; set; }
        public string Com_Country { get; set; }
        public string Trd_Curr { get; set; }
        public decimal? Curr_Out_Shares { get; set; }
        public string IMG_URL { get; set; }
        public string Comp_Country { get; set; }
        public string Sec_Name { get; set; }
        public string Ind_Name { get; set; }
    }
}