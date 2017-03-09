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
            log.Info("index!111");
            return View();
        }
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase  file,FormCollection form)
        {
            log.Info("index for Upload File:" + file.FileName);
            log.Info("get projectId=" + form.Get("projectId"));
            log.Info("get startrow=" + form.Get("startrow"));

            if (file.ContentLength > 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/UploadFile/TEST"), fileName);
                file.SaveAs(path);
                //Sample Code : 解析Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                poiservice.ConvertDataForTenderProject(form.Get("projectId"), int.Parse(form.Get("startrow")));
                log.Info("convert finish:" + poiservice.lstProjectItem.Count);
                ViewBag.result = "共" + poiservice.lstProjectItem.Count + "筆資料<br/>" + poiservice.errorMessage;
            }
            return View();
        }
    }
}