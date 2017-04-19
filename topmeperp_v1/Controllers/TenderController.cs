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
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName("", "備標");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View(lstProject);
        }
        // POST : Search
        public ActionResult Search()
        {
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName(Request["textProejctName"], "備標");
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
        public ActionResult Task(string id)
        {
            logger.Info("task assign page!!");
            ViewBag.projectid = id;
            TnderProject service = new TnderProject();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            SYS_USER u = (SYS_USER)Session["user"];
            ViewBag.createid = u.USER_ID;
            return View();
        }

        //POST:TaskAssign
        [HttpPost]
        public ActionResult Task(string id, List<TND_TASKASSIGN> TaskDatas)
        {
            logger.Info("task :" + Request["TaskDatas.index"]);
            TnderProject service = new TnderProject();
            service.refreshTask(TaskDatas);
            ViewBag.projectid = id;
            return View();
        }

        public ActionResult Details(string id)
        {
            logger.Info("project detail page projectid = " + id);
            List<TND_TASKASSIGN> lstTask = null;
            TnderProject service = new TnderProject();
            lstTask = service.getTaskById(id);
            TND_PROJECT p = service.getProjectById(id);
            TndProjectModels viewModel = new TndProjectModels();
            viewModel.tndProject = p;
            viewModel.tndTaskAssign = lstTask;
            return View(viewModel);
        }
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
            logger.Info("Upload Map Info for projectid=" + projectid);
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
                List<TND_MAP_DEVICE> lstMapDEVICE = poiservice.ConvertDataForMapDEVICE(projectid);
                //2.3 記錄錯誤訊息
                message = message + poiservice.errorMessage;
                //2.4
                logger.Info("Delete TND_MAP_DEVICE By Project ID");
                service.delMapDEVICEByProject(projectid);
                //2.5
                logger.Info("Add All TND_MAP_DEVICE to DB");
                service.refreshMapDEVICE(lstMapDEVICE);
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
            TempData["result"] = message;
            return RedirectToAction("MapInfoMainPage/" + projectid);

        }
        //建立各類別圖算數量檢視頁面
        public ActionResult MapInfoMainPage(string id)
        {
            logger.Info("mapinfo by projectID=" + id);
            ViewBag.projectid = id;
            string projectid = Request["projectid"];
            List<TND_MAP_FP> lstFP = null;
            List<TND_MAP_FW> lstFW = null;
            List<TND_MAP_PEP> lstPEP = null;
            List<TND_MAP_PLU> lstPLU = null;
            List<TND_MAP_LCP> lstLCP = null;
            List<TND_MAP_DEVICE> lstDEVICE = null;
            if (null != id && id != "")
            {
                TnderProject service = new TnderProject();
                lstFP = service.getMapFPById(id);
                lstFW = service.getMapFWById(id);
                lstPEP = service.getMapPEPById(id);
                lstPLU = service.getMapPLUById(id);
                lstLCP = service.getMapLCPById(id);
                lstDEVICE = service.getMapDEVICEById(id);
                MapInfoModels viewModel = new MapInfoModels();
                viewModel.mapFP = lstFP;
                viewModel.mapFW = lstFW;
                viewModel.mapPEP = lstPEP;
                viewModel.mapPLU = lstPLU;
                viewModel.mapLCP = lstLCP;
                viewModel.mapDEVICE = lstDEVICE;
                ViewBag.Result1 = "共有" + lstDEVICE.Count + "筆資料";
                ViewBag.Result2 = "共有" + lstFP.Count + "筆資料";
                ViewBag.Result3 = "共有" + lstFW.Count + "筆資料";
                ViewBag.Result4 = "共有" + lstPEP.Count + "筆資料";
                ViewBag.Result5 = "共有" + lstLCP.Count + "筆資料";
                ViewBag.Result6 = "共有" + lstPLU.Count + "筆資料";
                return View(viewModel);
            }
            return RedirectToAction("MapInfoMainPage/" + projectid);
        }
        //編輯消防電圖算數量
        #region 消防電
        public ActionResult EditForFP(string id)
        {
            logger.Info("get map fp for update:fp_id=" + id);
            TND_MAP_FP fp = null;
            if (null != id)
            {
                TnderProject service = new TnderProject();
                fp = service.getFPById(id);
            }
            return View(fp);
        }
        [HttpPost]
        public ActionResult EditForFP(TND_MAP_FP mapfp)
        {
            logger.Info("update map fp:fp_id=" + mapfp.FP_ID + "," + mapfp.EXCEL_ITEM);
            string message = "";
            TnderProject service = new TnderProject();
            service.updateMapFP(mapfp);
            message = "消防電圖算資料修改成功，ID :" + mapfp.FP_ID;
            ViewBag.result = message;
            return View(mapfp);
        }
        #endregion
        //編輯消防水圖算數量
        #region 消防水
        public ActionResult EditForFW(string id)
        {
            logger.Info("get map fw for update:fp_id=" + id);
            TND_MAP_FW fw = null;
            if (null != id)
            {
                TnderProject service = new TnderProject();
                fw = service.getFWById(id);
            }
            return View(fw);
        }
        [HttpPost]
        public ActionResult EditForFW(TND_MAP_FW mapfw)
        {
            logger.Info("update map fw:fw_id=" + mapfw.FW_ID + "," + mapfw.EXCEL_ITEM);
            string message = "";
            TnderProject service = new TnderProject();
            service.updateMapFW(mapfw);
            message = "消防水圖算資料修改成功，ID :" + mapfw.FW_ID;
            ViewBag.result = message;
            return View(mapfw);
        }
        #endregion
        //編輯電氣管線圖算數量
        #region 電氣管線
        public ActionResult EditForPEP(string id)
        {
            logger.Info("get mappep for update:pep_id=" + id);
            TND_MAP_PEP pep = null;
            if (null != id)
            {
                TnderProject service = new TnderProject();
                pep = service.getPEPById(id);
            }
            return View(pep);
        }
        [HttpPost]
        public ActionResult EditForPEP(TND_MAP_PEP mappep)
        {
            logger.Info("update map pep:pep_id=" + mappep.PEP_ID + "," + mappep.EXCEL_ITEM);
            string message = "";
            TnderProject service = new TnderProject();
            service.updateMapPEP(mappep);
            message = "電氣管線圖算資料修改成功，ID :" + mappep.PEP_ID;
            ViewBag.result = message;
            return View(mappep);
        }
        #endregion
        //編輯弱電管線圖算數量
        #region 弱電管線
        public ActionResult EditForLCP(string id)
        {
            logger.Info("get maplcp for update:lcp_id=" + id);
            TND_MAP_LCP lcp = null;
            if (null != id)
            {
                TnderProject service = new TnderProject();
                lcp = service.getLCPById(id);
            }
            return View(lcp);
        }
        [HttpPost]
        public ActionResult EditForLCP(TND_MAP_LCP maplcp)
        {
            logger.Info("update map lcp:lcp_id=" + maplcp.LCP_ID + "," + maplcp.EXCEL_ITEM);
            string message = "";
            TnderProject service = new TnderProject();
            service.updateMapLCP(maplcp);
            message = "弱電管線圖算資料修改成功，ID :" + maplcp.LCP_ID;
            ViewBag.result = message;
            return View(maplcp);
        }
        #endregion
        //編輯給排水圖算數量
        #region 給排水
        public ActionResult EditForPLU(string id)
        {
            logger.Info("get mapplu for update:plu_id=" + id);
            TND_MAP_PLU plu = null;
            if (null != id)
            {
                TnderProject service = new TnderProject();
                plu = service.getPLUById(id);
            }
            return View(plu);
        }
        [HttpPost]
        public ActionResult EditForPLU(TND_MAP_PLU mapplu)
        {
            logger.Info("update map plu:plu_id=" + mapplu.PLU_ID + "," + mapplu.EXCEL_ITEM);
            string message = "";
            TnderProject service = new TnderProject();
            service.updateMapPLU(mapplu);
            message = "給排水圖算資料修改成功，ID :" + mapplu.PLU_ID;
            ViewBag.result = message;
            return View(mapplu);
        }
        #endregion
        private List<topmeperp.Models.TND_PROJECT> SearchProjectByName(string projectname, string status)
        {
            if (projectname != null)
            {
                logger.Info("search project by 名稱 =" + projectname);
                List<topmeperp.Models.TND_PROJECT> lstProject = new List<TND_PROJECT>();
                using (var context = new topmepEntities())
                {
                    lstProject = context.TND_PROJECT.SqlQuery("select * from TND_PROJECT p "
                        + "where p.PROJECT_NAME Like '%' + @projectname + '%' AND STATUS=@status;",
                         new SqlParameter("projectname", projectname), new SqlParameter("status", status)).ToList();
                }
                logger.Info("get project count=" + lstProject.Count);
                return lstProject;
            }
            else
            {
                return null;
            }
        }

        public ActionResult EditProjectItem()
        {
            return View();
        }
    }
}