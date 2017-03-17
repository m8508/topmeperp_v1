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
            return View();
        }
        public ActionResult FunctionList()
        {
            string roleid = Request["roles"];
            log.Debug(Request.IsAjaxRequest());
            log.Info("index roleid=" + roleid);
            return View(userService.getFunctions(roleid));
        }
        public void UpdatePrivilege()
        {
            log.Info("new privilege:" + Request["hadPrivilege"]);
        }
    }
}