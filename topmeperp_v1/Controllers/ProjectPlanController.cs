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
    public class ProjectPlanController : Controller
    {
        static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        // GET: ProjectPaln
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult uploadFile(HttpPostedFileBase file)
        {
            //設置施工管理資料夾
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
            return View("Index");
        }
    }
}