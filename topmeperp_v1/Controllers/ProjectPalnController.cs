using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace topmeperp.Controllers
{
    public class ProjectPalnController : Controller
    {
        static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // GET: ProjectPaln
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult uploadFile(HttpPostedFileBase file)
        {
            log.Info("upload file!!");
            return View("Index");
        }
    }
}