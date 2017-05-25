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
                log.Debug("get project task by project" + Request["projectid"]);
                DataTable dt = planService.getProjectTask(Request["projectid"]);
                string htmlString = "<table class='table table-bordered'><tr>";
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    log.Debug("column name=" + dt.Columns[i].ColumnName);
                    htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                }
                htmlString = htmlString + "</tr>";
                foreach (DataRow dr in dt.Rows)
                {
                    htmlString = htmlString + "<tr>";
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        htmlString = htmlString + "<td>" + dr[i] + "</td>";
                    }
                    htmlString = htmlString + "</tr>";
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