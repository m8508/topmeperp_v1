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
        public ActionResult Query(FormCollection form)
        {
            log.Info("criteria user_id=" + form.Get("userid")+ ",username=" + form.Get("username")+",tel="+ form.Get("tel") +",roleid="+ form.Get("roles"));
            //由TempData 讀取資料
            //IEnumerable<SYS_ROLE> u = (IEnumerable <SYS_ROLE>) TempData["roles"];
            //log.Info("temp data=" + u.ToString());
            //將資料存入TempData 減少不斷讀取資料庫//Tempdata 僅保證一個Request,所以需移除再加入
            //TempData.Remove("roles");
            //TempData.Add("roles",u);
            //SelectList roles = new SelectList(u, "ROLE_ID", "ROLE_NAME");
            //ViewBag.roles = roles;
            //查詢使用者明細資料
            SYS_USER u_user = new SYS_USER();
            u_user.USER_ID = form.Get("userid");
            u_user.USER_NAME = form.Get("username");
            u_user.TEL = form.Get("tel");
            userService.getUserByCriteria(u_user, form.Get("roles").Trim());
            ViewBag.SearchResult = "User 共" + userService.userManageModels.sysUsers.Count() + "筆資料!!";
            //回傳部分網頁
            return PartialView("UserList", userService.userManageModels);
            //return View(userService.userManageModels);
        } 
        //新增或修改使用者
        public String addUser(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            //懶得把Form綁SYS_USER 直接先把Form 值填滿
            SYS_USER u = new SYS_USER();
            u.USER_ID = form.Get("u_userid").Trim();
            u.USER_NAME = form.Get("u_name").Trim();
            u.PASSWORD = form.Get("u_password").Trim();
            u.EMAIL= form.Get("u_email").Trim();
            u.TEL = form.Get("u_tel").Trim();
            u.TEL_EXT = form.Get("u_tel_ext").Trim();
            u.MOBILE = form.Get("u_mobile").Trim();
            u.ROLE_ID = form.Get("roles").Trim();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            u.CREATE_ID = loginUser.USER_ID;
            u.CREATE_DATE = DateTime.Now;
           int i = userService.addNewUser(u);
            if (i == 0)
            {
                msg = userService.message;
            }else
            {
                msg = "帳號更新成功";
            }

            log.Info("Request:user_ID=" + form["u_userid"]);
            return msg;
        } 
        //取得某一User 基本資料
        public string getUser(string userid)
        {
            log.Info("get user id=" + userid);
            SYS_USER u = userService.getUser(userid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string userJson = objSerializer.Serialize(u);
            log.Info("user info=" + userJson);
            return userJson;
        }  
    }
}