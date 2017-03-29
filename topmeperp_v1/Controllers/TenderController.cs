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
                p = service.getProjectById(id);
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
            if (prj.PROJECT_ID == "" || prj.PROJECT_ID == null)
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
                var path = Path.Combine(ContextService.strUploadPath + "/" + prj.PROJECT_ID, fileName);
                logger.Info("save excel file:" + path);
                file.SaveAs(path);
                //2.2 解析Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                poiservice.ConvertDataForTenderProject(prj.PROJECT_ID, (int)prj.START_ROW_NO);
                //2.3 記錄錯誤訊息
                message = message + "標單品項:共" + poiservice.lstProjectItem.Count + "筆資料<br/>" + poiservice.errorMessage;
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
        [HttpPost]
        public ActionResult Task(string id, TND_TASKASSIGN task)
        {
            logger.Info("task assign result page and project id =" + id);
            TnderProject service = new TnderProject();
            task.PROJECT_ID = id;
            service.newTask(task);
            return View(task);
        }

        //[HttpGet]
        //public ActionResult Details(string id)
        //{
        //    logger.Info("project detail page projectid = " + id);
        //    TnderProject service = new TnderProject();
        //    TND_PROJECT p = null;// service.getProjectById(id);
        //    return View(p);
        //}
        public ActionResult uploadMapInfo(string id)
        {
            logger.Info("upload map info for projectid=" + id);
            ViewBag.projectid = id;
            return View();
        }
        [HttpPost]
        public ActionResult uploadMapInfo(HttpPostedFileBase fileDevice, HttpPostedFileBase fileFP,
            HttpPostedFileBase fileFW, HttpPostedFileBase fileLCP,
            HttpPostedFileBase filePEP, HttpPostedFileBase filePLU)
        {
            string projectid = Request["projectid"];
            TnderProject service = new TnderProject();
            logger.Info("Upload Map Info  for projectid=" + projectid);
            string message = "";
            //檔案變數名稱需要與前端畫面對應
            #region 設備清單
            //圖算:設備檔案清單
            if (null != fileDevice && fileDevice.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + fileDevice.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileDevice.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                fileDevice.SaveAs(path);
                //2.2 解析Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                //解析設備圖算數量檔案
                //poiservice.ConvertDataForTenderProject(prj.PROJECT_ID, (int)prj.START_ROW_NO);
                //2.3 記錄錯誤訊息
                // message = message + "標單品項:共" + poiservice.lstProjectItem.Count + "筆資料<br/>" + poiservice.errorMessage;
                //2.4
                logger.Info("Delete TND_MAP_DEVICE By Project ID");
                //service.delAllItemByProject();
                //2.5
                logger.Info("Add All TND_MAP_DEVICE to DB");
                //service.refreshProjectItem(poiservice.lstProjectItem);
            }
            #endregion
            #region 消防電
            //圖算:消防電(TND_MAP_FP)
            if (null != fileFP && fileFP.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser FP Excel data:" + fileFP.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileFP.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                fileFP.SaveAs(path);
                //2.2 解析Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                //解析消防電圖算數量檔案
                List<TND_MAP_FP> lstMapFP = poiservice.ConvertDataForMapFP(projectid);
                //2.3 記錄錯誤訊息
                message = message + poiservice.errorMessage;
                //2.4
                logger.Info("Delete TND_MAP_FP By Project ID");
                service.delMapFPByProject(projectid);
                //2.5
                logger.Info("Add All TND_MAP_FP to DB");
                service.refreshMapFP(lstMapFP);
            }
            #endregion
            #region 消防水
            //圖算:消防水(TND_MAP_FW)
            if (null != fileFW && fileFW.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + fileFW.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileFW.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                fileFW.SaveAs(path);
                //2.2 開啟Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                //解析消防水塗算數量
                List<TND_MAP_FW> lstMapFW = poiservice.ConvertDataForMapFW(projectid);
                //2.3 記錄錯誤訊息
                message = poiservice.errorMessage;
                //2.4
                logger.Info("Delete TND_MAP_FW By Project ID");
                service.delMapFWByProject(projectid);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                logger.Info("Add All TND_MAP_FP to DB");
                service.refreshMapFW(lstMapFW);
                message = message + "<br/>資料匯入完成 !!";
            }
            #endregion
            #region 給排水
            //圖算:給排水(TND_MAP_PLU)
            if (null != filePLU && filePLU.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + filePLU.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(filePLU.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                filePLU.SaveAs(path);
                //2.2 開啟Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                //解析給排水圖算數量
                List<TND_MAP_PLU> lstMapPLU = poiservice.ConvertDataForMapPLU(projectid);
                //2.3 記錄錯誤訊息
                message = poiservice.errorMessage;
                //2.4
                logger.Info("Delete TND_MAP_PLU By Project ID");
                service.delMapPLUByProject(projectid);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                logger.Info("Add All TND_MAP_PLU to DB");
                service.refreshMapPLU(lstMapPLU);
                message = message + "<br/>資料匯入完成 !!";
            }
            #endregion
            #region 弱電管線
            //圖算:弱電管線(TND_MAP_LCP)
            if (null != fileLCP && fileLCP.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + fileLCP.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileLCP.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                fileLCP.SaveAs(path);
                //2.2 開啟Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                //解析弱電管線圖算數量
                List<TND_MAP_LCP> lstMapLCP = poiservice.ConvertDataForMapLCP(projectid);
                //2.3 記錄錯誤訊息
                message = poiservice.errorMessage;
                //2.4
                logger.Info("Delete TND_MAP_LCP By Project ID");
                service.delMapLCPByProject(projectid);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                logger.Info("Add All TND_MAP_LCP to DB");
                service.refreshMapLCP(lstMapLCP);
                message = message + "<br/>資料匯入完成 !!";
            }
            #endregion
            #region 電氣管線
            //圖算:電氣管線(TND_MAP_PEP)
            if (null != filePEP && filePEP.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + filePEP.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(filePEP.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                filePEP.SaveAs(path);
                //2.2 開啟Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                //解析電氣管線圖算數量
                List<TND_MAP_PEP> lstMapPEP = poiservice.ConvertDataForMapPEP(projectid);
                //2.3 記錄錯誤訊息
                message = poiservice.errorMessage;
                //2.4
                logger.Info("Delete TND_MAP_PEP By Project ID");
                service.delMapPEPByProject(projectid);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                logger.Info("Add All TND_MAP_PEP to DB");
                service.refreshMapPEP(lstMapPEP);
                message = message + "<br/>資料匯入完成 !!";
            }
            #endregion
            ViewBag.result = message;
            return RedirectToAction("MapInfoMainPage/" + projectid);
        }

        public ActionResult MapInfoMainPage(string id)
        {
            string projectid = Request["projectid"];
            logger.Info("mapinfo by projectID=" + id);
            List<TND_MAP_FP> lstFP = null;
            List<TND_MAP_FW> lstFW = null;
            if (null != id && id != "")
            {
                TnderProject service = new TnderProject();
                lstFP = service.getMapFPById(id);
                lstFW = service.getMapFWById(id);
                MapInfoModels viewModel = new MapInfoModels();
                viewModel.mapFP = lstFP;
                viewModel.mapFW = lstFW;

                return View(viewModel);
            }
            return RedirectToAction("MapInfoMainPage/" + projectid);
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