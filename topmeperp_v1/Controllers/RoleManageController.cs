using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    public class RoleManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(UserManageController));
        UserManage userService = new UserManage();
        // GET: RoleManage
        public ActionResult Index()
        {
            log.Info("index");
            userService.getAllRole();
            SelectList roles = new SelectList(userService.userManageModels.sysRole, "ROLE_ID", "ROLE_NAME");
            ViewBag.roles = roles;
            return View(userService.getPrivilege(""));
        }
        public ActionResult FunctionList()
        {
            string roleid = Request["roles"];
            log.Debug(Request.IsAjaxRequest());
            log.Info("index roleid=" + roleid);
            return PartialView(userService.getPrivilege(roleid));
        }
        public string UpdatePrivilege()
        {
            log.Info("new privilege:" + Request["hadPrivilege"]);
            return "還沒實作更新完成!!";
        }
    }
}