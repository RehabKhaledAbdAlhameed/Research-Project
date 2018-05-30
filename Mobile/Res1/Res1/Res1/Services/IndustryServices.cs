using Res1.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Res1.Services
{
    public class IndustryServices
    {

        public static List<Industry> GetIndustries()
        {
            var result = new List<Industry>() {

               new Industry(){Ind_Name="I1",Ind_id=1, Sec_id=1,IMG_URL=""},
               new Industry(){Ind_Name="I2",Ind_id=2, Sec_id=1,IMG_URL=""},
               new Industry(){Ind_Name="I3",Ind_id=3, Sec_id=2,IMG_URL=""},
               new Industry(){Ind_Name="I4",Ind_id=4, Sec_id=3,IMG_URL=""},
               new Industry(){Ind_Name="I5",Ind_id=5, Sec_id=2,IMG_URL=""},



            };
            return result;
        }
    }
}
