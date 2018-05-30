using Res1.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Res1.Services
{
    public class CompanyServices
    {

        public static List<Company> GetCompanies()
        {
            var result = new List<Company>() {

                new Company(){Comp_id=1,Comp_Name="ITWorx",IMG_URL="",Curr_Out_Shares=50,Ind_id=1,Sec_id=1,Reu_Code="ABC1",Com_Country="Egypt"},
                new Company(){Comp_id=2,Comp_Name="Microsoft",IMG_URL="",Curr_Out_Shares=50,Ind_id=1,Sec_id=1,Reu_Code="ABC2",Com_Country="Lebnon"},
                new Company(){Comp_id=3,Comp_Name="Google",IMG_URL="",Curr_Out_Shares=50,Ind_id=1,Sec_id=1,Reu_Code="ABC3",Com_Country="UAE"},
                new Company(){Comp_id=4,Comp_Name="IBM",IMG_URL="",Curr_Out_Shares=50,Ind_id=1,Sec_id=1,Reu_Code="ABC4",Com_Country="French"},
                new Company(){Comp_id=5,Comp_Name="Eye",IMG_URL="",Curr_Out_Shares=50,Ind_id=1,Sec_id=1,Reu_Code="ABC5",Com_Country="Germany"}




            };
            return result;
        } 
    }
}
