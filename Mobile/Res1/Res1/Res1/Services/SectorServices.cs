using Res1.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Res1.Services
{
    public class SectorServices
    {
        public static List<Sector> GetSectorss()
        {
            var result = new List<Sector>() {
        
               new Sector(){ Sec_id=1,Sec_Desc="Sector1",Sec_Name="Energy",IMG_URL="http://lorempixel.com/100/100/people/2"},
               new Sector(){ Sec_id=2,Sec_Desc="Sector2",Sec_Name="HealthCare",IMG_URL="http://lorempixel.com/100/100/people/2"},
               new Sector(){ Sec_id=3,Sec_Desc="Sector3",Sec_Name="RealEstate",IMG_URL="http://lorempixel.com/100/100/people/2"},
               new Sector(){ Sec_id=4,Sec_Desc="Sector4",Sec_Name="Utilities",IMG_URL="http://lorempixel.com/100/100/people/2"},
               new Sector(){ Sec_id=5,Sec_Desc="Sector5",Sec_Name="Industrials",IMG_URL="http://lorempixel.com/100/100/people/2"},



            };
            return result;
        }
    }
}
