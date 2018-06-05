using finalProj.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Cors;

namespace finalProj
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            //To enable any client access
            //First:setup (Microsoft.AspNet.WebApi.Cors) from Nuget
            var cors = new EnableCorsAttribute("*", "*", "*");
            config.EnableCors(cors);


            config.Formatters.Remove(config.Formatters.XmlFormatter);
            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
            name: "Api5",
            routeTemplate: "api/Industry_Specifec_Companies/{id}"
            );

            config.Routes.MapHttpRoute(
           name: "Api4",
           routeTemplate: "api/SectorIndustries/{id}"
           );


            config.Routes.MapHttpRoute(
            name: "Api3",
            routeTemplate: "api/compid/{id}"
            );


            config.Routes.MapHttpRoute(
            name: "Api2",
            routeTemplate: "api/IndustryCompanies/"
             );

            config.Routes.MapHttpRoute(
            name: "Api1",
            routeTemplate: "api/SectorCompanies/"
             );

            config.Routes.MapHttpRoute(
            name: "DefaultApi",
            routeTemplate: "api/{controller}/{id}",
            defaults: new { id = RouteParameter.Optional }
            );

           
        }

    }
}
