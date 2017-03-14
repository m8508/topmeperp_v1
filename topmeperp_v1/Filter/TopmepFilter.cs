using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace topmeperp.Filter
{
    //public class TopmepAuthorizeAttribute : AuthorizeAttribute
    //{
    //    ILog log = log4net.LogManager.GetLogger(typeof(TopmepAuthorizeAttribute));
    //    public override void OnAuthorization(AuthorizationContext filterContext)
    //    {
    //        log.Info("in OnAuthorization !!");
    //        //if (filterContext.ActionDescriptor.IsDefined(typeof(AllowAnonymousAttribute), true)
    //        //    || filterContext.ActionDescriptor.ControllerDescriptor.IsDefined(typeof(AllowAnonymousAttribute), true))
    //        //{
    //        //    // Don't check for authorization as AllowAnonymous filter is applied to the action or controller
    //        //    return;
    //        //}

    //        //// Check for authorization
    //        if (HttpContext.Current.Session["user"] == null)
    //        {
    //            log.Info("session user is not exist");
    //            filterContext.Result = new RedirectResult("/Home/Index");
    //            //base.HandleUnauthorizedRequest(filterContext);
    //            //filterContext.Result = filterContext.Result = new HttpUnauthorizedResult();
    //        }
    //        else
    //        {
    //            log.Info("session user is exist");
    //        }

    //    }
    //    //public override void OnAuthorization(AuthorizationContext filterContext)
    //    //{
    //    //    log.Info("in OnAuthorization !!");
    //    //    string token = "";
    //    //    if (filterContext.HttpContext.Session["user"] != null)
    //    //    {
    //    //        //驗證成功
    //    //        //filterContext.Result = new ViewResult() { ViewName = "Backend" };
    //    //        var loginTime = Convert.ToDateTime(filterContext.HttpContext.Application[token]);
    //    //        log.Info("session exist!!");
    //    //    }
    //    //    else
    //    //    {
    //    //        log.Info("session not exist!!");
    //    //        //base.HandleUnauthorizedRequest(filterContext);
    //    //    }
    //    //}

    //    protected override bool AuthorizeCore(HttpContextBase httpContext)
    //    {
    //        log.Info("in AuthorizeCore!!");
    //        return true;
    //        //if (httpContext == null)
    //        //{
    //        //    throw new ArgumentNullException("httpContext");
    //        //}

    //        //bool boolIsRight = false;
    //        //log.Info("check session!!");
    //        ////Session過期，要求重新登入
    //        //HttpSessionStateBase session = httpContext.Session;
    //        //if (session["user"] != null)
    //        //{
    //        //    boolIsRight = true;
    //        //}
    //        //log.Info("login result:" + boolIsRight);
    //        //return boolIsRight;
    //    }
    //}
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
                //if (context.HttpContext.Request.RawUrl == "/")
                //{
                //    log.Info("forward to first function");
                //}
            }
            else
            {
                log.Info("session not exist!!");
                if (context.HttpContext.Request.RawUrl != "/")
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