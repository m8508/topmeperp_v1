using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace topmeperp.Filter
{
    public class AuthFilter : ActionFilterAttribute
    {
        ILog log = log4net.LogManager.GetLogger(typeof(AuthFilter));
        public override void OnActionExecuting(ActionExecutingContext context)
        {
            log.Info("request URL=" + context.HttpContext.Request.RawUrl);
            if (context.HttpContext.Session["user"] != null)
            {
                //驗證成功
                log.Info("session exist!!");
            }
            else
            {
                log.Info("session not exist!!");
                if (context.HttpContext.Request.RawUrl != "/" 
                    && context.HttpContext.Request.RawUrl !="/topmeperp/")
                {
                    log.Info("forward to login page");
                    context.Result = new RedirectToRouteResult(
                    new RouteValueDictionary
                    {
                         { "controller", "Home" },
                         { "action", "Login" }
                    });
                }
            }

        }
    }

}