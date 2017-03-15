using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    public class UserManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(UserManageController));
        UserManage userService = new UserManage();
        // GET: UserManage
        public ActionResult Index()
        {
            log.Info("index");
            userService.getAllRole();
            ViewData.Add("roles", userService.userManageModels.sysRole);
            SelectList roles = new SelectList(userService.userManageModels.sysRole, "ROLE_ID", "ROLE_NAME");
            ViewBag.roles = roles;
            //將資料存入TempData 減少不斷讀取資料庫
            TempData.Remove("roles");
            TempData.Add("roles", userService.userManageModels.sysRole);      
            return View();
        }
        [HttpPost]
        public ActionResult Index(FormCollection form)
        {
            log.Info("criteria user_id=" + form.Get("userid")+ ",username=" + form.Get("username")+",tel="+ form.Get("tel") +",roleid="+ form.Get("roles"));
            //由TempData 讀取資料
            IEnumerable<SYS_ROLE> u = (IEnumerable <SYS_ROLE>) TempData["roles"];
            log.Info("temp data=" + u.ToString());
            //將資料存入TempData 減少不斷讀取資料庫//Tempdata 僅保證一個Request,所以需移除再加入
            TempData.Remove("roles");
            TempData.Add("roles",u);
            SelectList roles = new SelectList(u, "ROLE_ID", "ROLE_NAME");
            ViewBag.roles = roles;
            //查詢使用者明細資料
            SYS_USER u_user = new SYS_USER();
            u_user.USER_ID = form.Get("userid");
            u_user.USER_NAME = form.Get("username");
            u_user.TEL = form.Get("tel");
            userService.getUserByCriteria(u_user, form.Get("roles"));
            ViewBag.SearchResult = "共" + userService.userManageModels.sysUsers.Count() + "筆資料!!";
            return View(userService.userManageModels);
        } 
        public String addUser(FormCollection form)
        {
            log.Info("form:" + form.Count);
            log.Info("Request:user_ID=" + form["u_userid"]);
            return "增加使用者";
        }    
    }
}