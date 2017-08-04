using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    /// <summary>
    /// 專案任務相關功能
    /// </summary>
    public class ProjectPlanController : Controller
    {
        static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        ProjectPlanService planService = new ProjectPlanService();
        // GET: ProjectPaln
        public ActionResult Index()
        {
            if (null != Request["projectid"])
            {
                string projectid = Request["projectid"];
                string prjuid = null;
                PLAN_TASK task = null;
                log.Debug("get project task by project:" + projectid + ",roottag=" + Request["roottag"]);
                if (null != projectid && "" != projectid)
                {
                    prjuid = Request["prjuid"];
                    log.Debug("get project task by child task by prj_uid:" + prjuid);
                }

                if (null != Request["roottag"] && "Y" == Request["roottag"])
                {
                    task = planService.getRootTask(projectid);
                    if (null != task)
                    {
                        log.Debug("task=" + task.PRJ_UID);
                        prjuid = task.PRJ_UID.ToString();
                    }
                }

                DataTable dt = null;
                if (null == prjuid || prjuid == "")
                {
                    //取得所有任務
                    dt = planService.getProjectTask(projectid);
                }
                else
                {
                    //取得所有子項任務
                    dt = planService.getChildTask(projectid, int.Parse(prjuid));
                }
                string htmlString = "<table class='table table-bordered'>";

                htmlString = htmlString + "<tr><th>層級</th><th>任務名稱</th><th>開始時間</th><th>完成時間</th><th>工期</th><th>--</th><th>--</th></tr>";
                foreach (DataRow dr in dt.Rows)
                {
                    DateTime stardate = DateTime.Parse(dr[4].ToString());
                    DateTime finishdate = DateTime.Parse(dr[5].ToString());

                    htmlString = htmlString + "<tr><td>" + dr[1] + "<input type='checkbox' name='roottag' id='roottag' onclick='setRootTask(" + dr[2] + ")' /></td><td>" + dr[0] + "</td>"
                        + "<td>" + stardate.ToString("yyyy-MM-dd") + "</td><td>" + finishdate.ToString("yyyy-MM-dd") + "</td><td>" + dr[6] + "</td>"
                        + "<td ><a href =\"Index?projectid=" + projectid + "&prjuid=" + dr[3] + "\">上一層 </a></td>"
                        + "<td><a href=\"Index?projectid=" + projectid + "&prjuid=" + dr[2] + "\">下一層 </a></td></tr>";
                }
                htmlString = htmlString + "</table>";
                ViewBag.htmlResult = htmlString;
                ViewBag.projectId = Request["projectid"];
            }
            return View();
        }
        //上傳project 檔案，建立專案任務
        public ActionResult uploadFile(HttpPostedFileBase file)
        {
            //設置施工管理資料夾
            if (null != file)
            {
                log.Info("upload file!!" + file.FileName);
                string projectId = Request["projectid"];
                string projectFolder = ContextService.strUploadPath + "/" + projectId + "/" + ContextService.projectMgrFolder;
                if (Directory.Exists(projectFolder))
                {
                    //資料夾存在
                    log.Info("Directory Exist:" + projectFolder);
                }
                else
                {
                    //if directory not exist create it
                    Directory.CreateDirectory(projectFolder);
                }
                if (null != file && file.ContentLength != 0)
                {
                    //2.upload project file
                    //2.2 將上傳檔案存檔
                    var fileName = Path.GetFileName(file.FileName);
                    var path = Path.Combine(projectFolder, fileName);
                    file.SaveAs(path);
                    OfficeProjectService s = new OfficeProjectService();
                    s.convertProject(projectId, path);
                    s.import2Table();
                }
            }
            return Redirect("Index?projectid=" + Request["projectid"] + "&roottag=" + Request["roottag"]);
            // return View("Index/projectid=" + Request["projectid"]);
        }
        //設定合約範圍起始任務
        public string setRootFlag()
        {
            log.Debug("projectid=" + Request["projectid"] + ",prjuid=" + Request["prjuid"]);
            int i = planService.setRootTask(Request["projectid"], Request["prjuid"]);
            return "設定完成!!(" + i + ")";
        }
        //專案任務與圖算數量設定畫面
        public ActionResult ManageTaskDetail()
        {
            log.Debug("show sreen for task manage");
            string projectid = Request["projectid"];
            ViewBag.projectId = projectid;
            ViewBag.TreeString = planService.getProjectTask4Tree(projectid); ;
            return View();
        }
        //查詢圖算資訊
        public ActionResult getMapItem4Task(FormCollection f)
        {
            string projectid = f["projectid"];
            log.Debug("projectid" + f["projectid"]);
            string mapno = f["mapno"];
            log.Debug("mapno" + f["mapno"]);
            string buildno = f["buildno"];
            log.Debug("buildno" + f["buildno"]);
            string primeside = f["primeside"];
            log.Debug("primeside" + f["primeside"]);
            string secondside = f["secondside"];
            log.Debug("secondside" + f["secondside"]);
            string devicename = f["devicename"];
            log.Debug("devicename" + f["devicename"]);
            string mapType = f["mapType"];
            log.Debug("mapType" + f["mapType"]);
            string strart_id = f["startid"];
            string end_id = f["endid"];
            if (null == f["mapType"] || "" == f["mapType"])
            {
                ViewBag.Message = "至少需選擇一項施作項目!!";
                return PartialView("_getMapItem4Task", null);
            }
            string[] mapTypes = mapType.Split(',');
            for (int i = 0; i < mapTypes.Length; i++)
            {
                switch (mapTypes[i])
                {
                    case "MAP_DEVICE"://設備
                        log.Debug("MapType: MAP_DEVICE(設備)");
                        planService.getMapItem(projectid, devicename, strart_id, end_id);
                        break;
                    case "MAP_PEP"://電器管線
                        log.Debug("MapType: MAP_PEP(電器管線)");
                        planService.getMapPEP(projectid, mapno, buildno, primeside, secondside, devicename);
                        break;
                    case "MAP_LCP"://弱電管線
                        log.Debug("MapType: MAP_LCP(弱電管線)");
                        planService.getMapLCP(projectid, mapno, buildno, primeside, secondside, devicename);
                        break;
                    case "TND_MAP_PLU"://給排水
                        log.Debug("MapType: TND_MAP_PLU(給排水)");
                        planService.getMapPLU(projectid, mapno, buildno, primeside, secondside, devicename);
                        break;
                    case "MAP_FP"://消防電
                        log.Debug("MapType: MAP_FP(消防電)");
                        planService.getMapFP(projectid, mapno, buildno, primeside, secondside, devicename);
                        break;
                    case "MAP_FW"://消防水
                        planService.getMapFW(projectid, mapno, buildno, primeside, secondside, devicename);
                        log.Debug("MapType: MAP_FW(消防水)");
                        break;
                    default:
                        log.Debug("MapType nothing!!");
                        break;
                }
            }
            ViewBag.Message = planService.resultMessage;
            return PartialView("_getMapItem4Task", planService.viewModel);
        }
        //設定任務圖算
        public string choiceMapItem(FormCollection f)
        {
            if (null == f["checkNodeId"] || "" == f["checkNodeId"])
            {
                return "請選擇專案任務!!";
            }
            if (null != f["map_device"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",mapids=" + f["map_device"]);
                //設備
                int i = planService.choiceMapItem(f["projectid"], f["checkNodeId"], f["map_device"]);
                log.Debug("modify records count=" + i);
            }
            if (null != f["map_pep"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_pep=" + f["map_pep"]);
                //電器管線
                int i = planService.choiceMapItemPEP(f["projectid"], f["checkNodeId"], f["map_pep"]);
                log.Debug("modify records count=" + i);
            }
            if (null != f["map_lcp"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_lcp=" + f["map_lcp"]);
                //弱電
                int i = planService.choiceMapItemLCP(f["projectid"], f["checkNodeId"], f["map_lcp"]);
                log.Debug("modify records count=" + i);
            }

            if (null != f["map_plu"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_plu=" + f["map_plu"]);
                //給排水
                int i = planService.choiceMapItemPLU(f["projectid"], f["checkNodeId"], f["map_plu"]);
                log.Debug("modify records count=" + i);
            }

            if (null != f["map_fp"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_fp=" + f["map_fp"]);
                //消防電
                int i = planService.choiceMapItemFP(f["projectid"], f["checkNodeId"], f["map_fp"]);
                log.Debug("modify records count=" + i);
            }
            if (null != f["map_fw"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_fw=" + f["map_fw"]);
                //消防水
                int i = planService.choiceMapItemFW(f["projectid"], f["checkNodeId"], f["map_fw"]);
                log.Debug("modify records count=" + i);
            }

            return "設定成功";
        }
        public ActionResult getActionItem4Task(FormCollection f)
        {
            log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"]);
            //planService
            return PartialView("_getProjecttem4Task", planService.getItemInTask(f["projectid"], f["checkNodeId"]));
        }
        //填寫日報step1 :選取任務
        public ActionResult dailyReport(string id)
        {
            if (null == id || "" == id)
            {
                id = Request["projectid"];
            }
            string strRptDate = "";
            if (null != Request["reportDate"])
            {
                strRptDate = Request["reportDate"].Trim();
            }
            DateTime dtTaskDate = DateTime.Now;
            if (strRptDate != "")
            {
                dtTaskDate = DateTime.Parse(strRptDate);
            }
            log.Debug("get Task for plan by day=" + dtTaskDate);
            List<PLAN_TASK> lstTask = planService.getTaskByDate(id, dtTaskDate);
            ViewBag.projectName = planService.getProject(id).PROJECT_NAME;
            ViewBag.projectId = id;
            ViewBag.reportDate = dtTaskDate.ToString("yyyy/MM/dd");
            return View(lstTask);
        }
        //填寫日報step2 :選取填寫內容
        public ActionResult dailyReportItem()
        {
            ViewBag.projectId = Request["projectid"];
            ViewBag.projectName = planService.getProject(Request["projectid"]).PROJECT_NAME;
            ViewBag.prj_uid = Request["prjuid"];
            ViewBag.taskName = planService.getProjectTask(Request["projectid"], int.Parse(Request["prjuid"])).TASK_NAME;
            ViewBag.RptDate = Request["rptDate"];
            //1.依據任務取得相關施作項目內容
            DailyReport dailyRpt = planService.newDailyReport(Request["projectid"], int.Parse(Request["prjuid"]));
            return View(dailyRpt);
        }
        //儲存日報數量紀錄
        public string saveItemRow(FormCollection f)
        {
            SYS_USER u = (SYS_USER)Session["user"];
            log.Debug("projectId=" + f["Projectid"] + ",prjUid=" + f["PrjUid"] + ",ReportId=" + f["ReportID"]);
            log.Debug("form Data ItemId=" + f["planItemId"]);
            log.Debug("form Data Qty=planItemQty" + f["planItemQty"]);


            string projectid = f["Projectid"];
            int prjuid = int.Parse(f["PrjUid"]);
            string strWeather = f["selWeather"];
            string strSummary = f["txtSummary"];
            string strSenceUser = f["txtSenceUser"];
            string strSupervision = f["txtSupervision"];
            string strOwner = f["txtOwner"];
            string strRptDate = f["RptDate"];

            DailyReport newDailyRpt = new DailyReport();
            PLAN_DALIY_REPORT RptHeader = new PLAN_DALIY_REPORT();
            RptHeader.PROJECT_ID = projectid;
            RptHeader.WEATHER = strWeather;
            RptHeader.SUMMARY = strSummary;
            RptHeader.SCENE_USER_NAME = strSenceUser;
            RptHeader.SUPERVISION_NAME = strSupervision;
            RptHeader.OWNER_NAME = strOwner;
            newDailyRpt.dailyRpt = RptHeader;
            RptHeader.REPORT_DATE = DateTime.Parse(strRptDate);
            //取得日報編號
            SerialKeyService snService = new SerialKeyService();
            if (null == f["ReportID"] || "" == f["ReportID"])
            {
                RptHeader.REPORT_ID = snService.getSerialKey(planService.KEY_ID);
            }else
            {
                RptHeader.REPORT_ID = f["ReportID"];
            }

            RptHeader.CREATE_DATE = DateTime.Now;
            RptHeader.CREATE_USER_ID = u.USER_ID;
            //建立專案任務資料 (結構是支援多項任務，僅先使用一筆)
            newDailyRpt.lstRptTask = new List<PLAN_DR_TASK>();
            PLAN_DR_TASK RptTask = new PLAN_DR_TASK();
            RptTask.PROJECT_ID = projectid;
            RptTask.PRJ_UID = prjuid;
            RptTask.REPORT_ID = RptHeader.REPORT_ID;
            newDailyRpt.lstRptTask.Add(RptTask);
            //處理料件
            newDailyRpt.lstRptItem = new List<PLAN_DR_ITEM>();
            string[] aryPlanItem = f["planItemId"].Split(',');
            string[] aryPlanItemQty = f["planItemQty"].Split(',');
            log.Debug("count ItemiD=" + aryPlanItem.Length + ",qty=" + aryPlanItemQty.Length);
            newDailyRpt.lstRptItem = new List<PLAN_DR_ITEM>();
            for (int i = 0; i < aryPlanItem.Length; i++)
            {
                PLAN_DR_ITEM item = new PLAN_DR_ITEM();
                item.PLAN_ITEM_ID = aryPlanItem[i];
                item.PROJECT_ID = projectid;
                item.REPORT_ID = RptHeader.REPORT_ID;
                if ("" != aryPlanItemQty[i])
                {
                    item.FINISH_QTY = decimal.Parse(aryPlanItemQty[i]);
                }
                newDailyRpt.lstRptItem.Add(item);
            }
            //處理出工資料
            newDailyRpt.lstRptWorkerAndMachine = new List<PLAN_DR_WORKER>();
            string[] aryWorkerType = f["workerKeyid"].Split(',');
            string[] aryWorkerQty = f["planWorkerQty"].Split(',');
            for (int i = 0; i < aryWorkerType.Length; i++)
            {
                PLAN_DR_WORKER item = new PLAN_DR_WORKER();
                item.REPORT_ID = RptHeader.REPORT_ID;
                item.WORKER_TYPE = "WORKER";
                item.PARA_KEY_ID = aryWorkerType[i];

                if ("" != aryWorkerQty[i].Trim())
                {
                    item.WORKER_QTY = decimal.Parse(aryWorkerQty[i]);
                    newDailyRpt.lstRptWorkerAndMachine.Add(item);
                }
            }
            log.Debug("count WorkerD=" + f["workerKeyid"] + ",WorkerQty=" + f["planWorkerQty"]);

            //處理機具資料
            string[] aryMachineType = f["MachineKeyid"].Split(',');
            string[] aryMachineQty = f["planMachineQty"].Split(',');
            for (int i = 0; i < aryMachineType.Length; i++)
            {
                PLAN_DR_WORKER item = new PLAN_DR_WORKER();
                item.REPORT_ID = RptHeader.REPORT_ID;
                item.WORKER_TYPE = "MACHINE";
                item.PARA_KEY_ID = aryMachineType[i];
                if ("" != aryMachineQty[i])
                {
                    item.WORKER_QTY = decimal.Parse(aryMachineQty[i]);
                    newDailyRpt.lstRptWorkerAndMachine.Add(item);
                }
            }
            log.Debug("count MachineD=" + f["MachineKeyid"] + ",WorkerQty=" + f["planMachineQty"]);
            //處理重要事項資料
            newDailyRpt.lstRptNote = new List<PLAN_DR_NOTE>();
            string[] aryNote = f["planNote"].Split(',');
            for (int i = 0; i < aryNote.Length; i++)
            {
                PLAN_DR_NOTE item = new PLAN_DR_NOTE();
                item.REPORT_ID = RptHeader.REPORT_ID;
                if ("" != aryNote[i].Trim())
                {
                    item.REMARK = aryNote[i].Trim();
                    newDailyRpt.lstRptNote.Add(item);
                }
            }
            log.Debug("count Note=" + f["planNote"]);
            string msg = planService.createDailyReport(newDailyRpt);
            return msg;
        }
        public ActionResult dailyReportList(string id)
        {
            if (null == id || "" == id)
            {
                id = Request["projectid"];
            }
            ViewBag.projectName = planService.getProject(id).PROJECT_NAME;
            ViewBag.projectId = id;
            return View();
        }
    }
}