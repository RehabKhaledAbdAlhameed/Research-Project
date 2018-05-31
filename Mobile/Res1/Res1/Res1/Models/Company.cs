using System;
using System.Collections.Generic;
using System.Text;

namespace Res1.Models
{
    public class Company
    {
       
        public int Comp_id { get; set; }

      
        public string Comp_Name { get; set; }

        public string Reu_Code { get; set; }

        public string Trd_Curr { get; set; }

        public decimal? Curr_Out_Shares { get; set; }

        public string IMG_URL { get; set; }

        public string Com_Country { get; set; }

        public int Sec_id { get; set; }

        public int Ind_id { get; set; }

        public virtual Industry Industry { get; set; }

        public virtual Sector Sector { get; set; }
    }
}
