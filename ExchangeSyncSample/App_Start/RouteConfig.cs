using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace ExchangeSyncSample
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            // Register the default hubs route: ~/signalr
            //RouteTable.Routes.MapHubs();
            routes.MapHubs();

            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Schedule", action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}