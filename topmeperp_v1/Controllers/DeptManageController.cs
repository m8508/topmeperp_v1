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
    public class DeptManageController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        UserManage userService = new UserManage();
        // GET: UserManage
        public ActionResult Index()
        {
            logger.Info("index");
            return View();
        }
        [HttpPost]
        public ActionResult Query(FormCollection form)
        {
            logger.Info("criteria user_id=" + form.Get("userid") + ",username=" + form.Get("username") + ",tel=" + form.Get("tel") + ",roleid=" + form.Get("roles"));
            return PartialView("UserList", userService.userManageModels);
            //return View(userService.userManageModels);
        }
        //新增或修改部門
        public String addUser(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            //懶得把Form綁SYS_USER 直接先把Form 值填滿
            SYS_USER u = new SYS_USER();
            u.USER_ID = form.Get("u_userid").Trim();
            u.USER_NAME = form.Get("u_name").Trim();
            u.PASSWORD = form.Get("u_password").Trim();
            u.EMAIL = form.Get("u_email").Trim();
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
            }
            else
            {
                msg = "帳號更新成功";
            }

            logger.Info("Request:user_ID=" + form["u_userid"]);
            return msg;
        }
        //取得某一Dept 基本資料
        public string getUser(string deptid)
        {
            logger.Info("get Dept id=" + deptid);

            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string userJson = objSerializer.Serialize("");
            logger.Info("user info=" + userJson);
            return userJson;
        }


    }
}