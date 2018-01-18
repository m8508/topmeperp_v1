using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using topmeperp.Models;

namespace topmeperp.Filter
{
    public class AuthFilter : ActionFilterAttribute
    {
        static ILog log = log4net.LogManager.GetLogger(typeof(AuthFilter));
        public override void OnActionExecuting(ActionExecutingContext context)
        {
            log.Info("request URL=" + context.HttpContext.Request.RawUrl);
            if (context.HttpContext.Session["user"] != null)
            {
                //驗證成功
                log.Info("session exist!!");
                //if (context.HttpContext.Request.RawUrl == "/")
                //{
                //    log.Info("forward to first function");
                //}
            }
            else
            {
                log.Info("session not exist!!");
                if (context.HttpContext.Request.RawUrl != "/"
                    && context.HttpContext.Request.RawUrl != "/topmeperp/")
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
        public static string getBreadcrumb(string url)
        {
            string breadcrumbHtml = "";
            List<SYS_FUNCTION> lst = (List<SYS_FUNCTION>)HttpContext.Current.Session["functions"];
            //使用list 物件查詢功能
            SYS_FUNCTION curFunction = lst.Find(x => x.FUNCTION_URI.Contains(url));
            log.Info("cur url=" + curFunction);

            if (curFunction.ISMENU == "Y")
            {
                //string[] breadcrumb = url.Split('/');
                // for (int i = 1; i < breadcrumb.Length; i++)
                // { breadcrumbHtml = breadcrumbHtml + "<li class='breadcrumb-item'><a href='"+ curFunction.FUNCTION_URI + "'>"+ curFunction.FUNCTION_NAME+ "</a></li>";
                breadcrumbHtml = breadcrumbHtml + "<li class='breadcrumb-item'><a href='#'>" + curFunction.MODULE_NAME + "</a></li>";
                breadcrumbHtml = breadcrumbHtml + "<li class='breadcrumb-item'><a href='" + curFunction.FUNCTION_URI + "'>" + curFunction.FUNCTION_NAME + "</a></li>";
                // }
            }
            return breadcrumbHtml;
        }
    }

}