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
        public ActionResult Create(String id)
        {
            logger.Info("get project for update:project_id=" + id);
            TND_PROJECT p = null;
            if (null != id)
            {
                TnderProject service = new TnderProject();
                p= service.getProjectById(id);
            }
            return View(p);
        }
        [HttpPost]
        public ActionResult Create(TND_PROJECT prj, HttpPostedFileBase file)
        {
            logger.Info("create project process! project =" + prj.ToString());
            TnderProject service = new TnderProject();
            SYS_USER u = (SYS_USER)Session["user"];
            string message = "";
            //1.更新或新增專案基本資料
            if (prj.PROJECT_ID=="" || prj.PROJECT_ID == null)
            {
                //新增專案
                prj.STATUS = "備標";
                prj.CREATE_USER_ID = u.USER_ID;
                prj.OWNER_USER_ID = u.USER_ID;
                prj.CREATE_DATE = DateTime.Now;
                service.newProject(prj);
                message = "建立專案:" + prj.PROJECT_ID + "<br/>";
            }
            else
            {
                //修改專案基本資料
                prj.MODIFY_USER_ID = u.USER_ID;
                prj.MODIFY_DATE = DateTime.Now;
                service.updateProject(prj);
                message = "專案基本資料修改:" + prj.PROJECT_ID + "<br/>";
            }
            //若使用者有上傳標單資料，則增加標單資料
            if (null != file && file.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + file.FileName);
                //2.1 設定Excel 檔案名稱
                prj.EXCEL_FILE_NAME = file.FileName;
                //2.2 將上傳檔案存檔
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(TnderProject.UploadFolder + "/" + prj.PROJECT_ID, fileName);
                logger.Info("save excel file:" + path);
                file.SaveAs(path);
                //2.2 解析Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                poiservice.ConvertDataForTenderProject(prj.PROJECT_ID, (int)prj.START_ROW_NO);
                //2.3 記錄錯誤訊息
                message = message +  "標單品項:共" + poiservice.lstProjectItem.Count + "筆資料<br/>" + poiservice.errorMessage;
                //2.4
                logger.Info("Delete TND_PROJECT_ITEM By Project ID");
                service.delAllItemByProject();
                //2.5
                logger.Info("Add All TND_PROJECT_ITEM to DB");
                service.refreshProjectItem(poiservice.lstProjectItem);
            }
            ViewBag.result = message;
            return View(prj);
        }


        //POST:TaskAssign
        public ActionResult Task()
        {
            logger.Info("task assign page!!");
            return View(); 
        }
        
        //public ActionResult RFQ(string id)
        //{
        //    logger.Info("redirect to Inquery page:projectid=" + id);
        //    return Redirect("/Inquiry/Index/" + id);
        //    //return RedirectToAction("Index", "Inquiry");
        //}

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

