using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
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
                log.Debug("get project task by project:" + projectid);
                if (null != Request["projectid"] && "" != Request["projectid"])
                {
                    prjuid = Request["prjuid"];
                    log.Debug("get project task by child task by prj_uid:" + prjuid);
                }

                DataTable dt = null;
                if (null == prjuid || prjuid=="")
                {
                    //取得所有任務
                    dt= planService.getProjectTask(projectid);
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

                    htmlString = htmlString + "<tr><td>" + dr[1] + "</td><td>" + dr[0] + "</td>"
                        + "<td>" + stardate.ToString("yyyy-MM-dd") + "</td><td>"+ finishdate.ToString("yyyy-MM-dd") + "</td><td>"+ dr[6] + "</td>"
                        + "<td ><a href =\"Index?projectid=" + projectid + "&prjuid=" + dr[3] + "\">上一層 </a></td>"
                        + "<td><a href=\"Index?projectid=" + projectid + "&prjuid=" + dr[2] + "\">下一層 </a></td></tr>";
                }
                htmlString = htmlString + "</table>";
                ViewBag.htmlResult = htmlString;
                ViewBag.projectId = Request["projectid"];
            }
            return View();
        }
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
            return Redirect("Index?projectid=" + Request["projectid"]);
            // return View("Index/projectid=" + Request["projectid"]);
        }
    }
}