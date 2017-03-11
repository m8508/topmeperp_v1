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
    public class InquiryController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(InquiryController));

        // GET: Inquiry
        public ActionResult Index()
        {
            log.Info("index!");
            return View();
        }
        // POST : Search
        public ActionResult Search()
        {
            List<topmeperp.Models.TND_PROJECT_ITEM> lstProject = SearchProjectByName(Request["textProejctItem"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View("Index", lstProject);
        }
        private List<topmeperp.Models.TND_PROJECT_ITEM> SearchProjectByName(string projectitem)
        {

            log.Info("search project by 名稱 =" + projectitem);
            List<topmeperp.Models.TND_PROJECT_ITEM> lstProject = new List<TND_PROJECT_ITEM>();
            using (var context = new topmepEntities())
            {
                lstProject = context.TND_PROJECT_ITEM.SqlQuery("select * from TND_PROJECT_ITEM p "
                    + "where p.TYPE_CODEE_1 Like '%' + @projectitem + '%';",
                     new SqlParameter("projectitem", projectitem)).ToList();
            }
            log.Info("get project count=" + lstProject.Count);
            return lstProject;
        }

    }
}

