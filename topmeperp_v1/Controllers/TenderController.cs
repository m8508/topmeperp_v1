using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;

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
        public ActionResult Create()
        {
            return View();
        }
    public ActionResult Edit()
        {
            return View();
        }

        private List<topmeperp.Models.TND_PROJECT> SearchProjectByName(string projectname)
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
    }
}