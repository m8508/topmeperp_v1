using log4net;
using System;
using System.Data;
using System.IO;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
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
                        planService.getMapItem(projectid, devicename);
                        break;
                    case "MAP_PEP"://電器管線
                        log.Debug("MapType: MAP_PEP(電器管線)");
                        break;
                    case "MAP_LCP"://弱電管線
                        log.Debug("MapType: MAP_LCP(弱電管線)");
                        break;
                    case "TND_MAP_PLU"://給排水
                        log.Debug("MapType: TND_MAP_PLU(給排水)");
                        planService.getMapPLU(projectid, mapno, buildno, primeside, secondside, devicename);
                        break;
                    case "MAP_FP"://消防電
                        log.Debug("MapType: MAP_FP(消防電)");
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
            if (null != f["map_fw"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_fw=" + f["map_fw"]);
                //消防水
                int i = planService.choiceMapItemFW(f["projectid"], f["checkNodeId"], f["map_fw"]);
                log.Debug("modify records count=" + i);
            }

            if (null != f["map_plu"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_fw=" + f["map_plu"]);
                //給排水
                int i = planService.choiceMapItemPLU(f["projectid"], f["checkNodeId"], f["map_plu"]);
                log.Debug("modify records count=" + i);
            }
            else
            {
                return "有問題!!";
            }
            return "設定成功";
        }
    }
}