using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;
using System.IO;

namespace topmeperp.Controllers
{
    public class TenderController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        // GET: Tender
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            return View();
        }
        // POST : Search
        public ActionResult Search()
        {
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName(Request["textProejctName"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View("Index", lstProject);
        }
        //POST:Create
        public ActionResult Create()
        {
            logger.Info("new project page!!");
            return View();
        }
        [HttpPost]
        public ActionResult Create(TND_PROJECT prj, HttpPostedFileBase file)
        {
            logger.Info("create project process! project name=" + prj.PROJECT_NAME + ",ProjectId=" + prj.PROJECT_ID );
            //1.更新或新增專案基本資料
            if (file.ContentLength != 0)
            {
                logger.Info("Parser Excel data:" +file.FileName);
                logger.Info("Delete TND_PROJECT_ITEM By Project ID");
                logger.Info("Add All TND_PROJECT_ITEM to DB");
            }
            //TnderProject service = new TnderProject();
            //service.newProject(prj);
            return View("~/Views/Tender/Index.cshtml");
        }
        //POST:TaskAssign
        public ActionResult Task()
        {
            logger.Info("task assign page!!");
            return View(); 
        }
        
        public ActionResult RFQ()
        {
            logger.Info("inquiry page!!");
            return RedirectToAction("Index", "Inquiry");
        }

        [HttpGet]
        public ActionResult Details(string id)
        {
            logger.Info("project detail page projectid = " + id);
            TnderProject service = new TnderProject();
            TND_PROJECT p = null;// service.getProjectById(id);
            return View(p);
        }
        [HttpGet]
        public ActionResult TaskDetails(string id)
        {
            logger.Info("taskassign detail page projectid = " + id);
            TnderProject service = new TnderProject();
            TND_TASKASSIGN t = service.getTaskById(id);
            return View(t);
        }
        public ActionResult Edit()
        {
            return View();
        }
        public ActionResult Upload()
        {
            logger.Info("upload!");
            return View();
        }
        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file, FormCollection form)
        {
            if (file != null)
            {
                logger.Info("index for Upload File:" + file.FileName);
                logger.Info("get projectId=" + form.Get("projectId"));
                logger.Info("get startrow=" + form.Get("startrow"));

                if (file.ContentLength > 0)
                {
                    var fileName = Path.GetFileName(file.FileName);
                    var path = Path.Combine(Server.MapPath("~/UploadFile/TEST"), fileName);
                    file.SaveAs(path);
                    //Sample Code : 解析Excel 檔案
                    ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                    poiservice.InitializeWorkbook(path);
                    poiservice.ConvertDataForTenderProject(form.Get("projectId"), int.Parse(form.Get("startrow")));
                    logger.Info("convert finish:" + poiservice.lstProjectItem.Count);
                    ViewBag.result = "共" + poiservice.lstProjectItem.Count + "筆資料<br/>" + poiservice.errorMessage;
                }
                return View("~/Views/Tender/Upload.cshtml");
            }
            else { return View(); }
        }
        private List<topmeperp.Models.TND_PROJECT> SearchProjectByName(string projectname)
        {
            if (projectname != null)
            {
                logger.Info("search project by 名稱 =" + projectname);
                List<topmeperp.Models.TND_PROJECT> lstProject = new List<TND_PROJECT>();
                using (var context = new topmepEntities())
                {
                    lstProject = context.TND_PROJECT.SqlQuery("select * from TND_PROJECT p "
                        + "where p.PROJECT_NAME Like '%' + @projectname + '%';",
                         new SqlParameter("projectname", projectname)).ToList();
                }
                logger.Info("get project count=" + lstProject.Count);
                return lstProject;
            }
            else
            {
                return null;
            }
        }

    }
}

