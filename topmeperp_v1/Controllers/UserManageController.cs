using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    public class UserManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(UserManageController));
        // GET: UserManage
        public ActionResult Index()
        {
            log.Info("index");
            return View();
        }
        [HttpPost]
        public ActionResult Index(FormCollection form)
        {
            log.Info("criteria user_id=" + form.Get("userid")+ ",username=" + form.Get("username")+",tel="+ form.Get("tel"));
            //查詢使用者明細資料
            return View();
        }
        
    }
}