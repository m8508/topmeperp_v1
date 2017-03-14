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
            log.Info("inquiry index!");
            return View();
        }

        // POST : Search
        public ActionResult Search()
        {
            List<topmeperp.Models.TND_PROJECT_ITEM> lstProject = SearchProjectByName(Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View("Index", lstProject);
        }
        private List<topmeperp.Models.TND_PROJECT_ITEM> SearchProjectByName(string projectitem1, string projectitem2, string projectitem3, string projectitem4)
        {

            log.Info("search projectitem by 九宮格 =" + projectitem1 + "search projectitem by 次九宮格 =" + projectitem2 + "search projectitem by 主系統 =" + projectitem3 + "search projectitem by 次系統 =" + projectitem4);
            List<topmeperp.Models.TND_PROJECT_ITEM> lstProject = new List<TND_PROJECT_ITEM>();
            using (var context = new topmepEntities())
            {
                lstProject = context.TND_PROJECT_ITEM.SqlQuery("select * from TND_PROJECT_ITEM p "
                    + "where p.TYPE_CODE_1 Like '%' + @projectitem1 + '%'and p.TYPE_CODE_2 Like '%' + @projectitem2 + '%'and p.SYSTEM_MAIN Like '%' + @projectitem3 + '%' and p.SYSTEM_SUB Like '%' + @projectitem4 + '%';",
                     new SqlParameter("projectitem1", projectitem1), new SqlParameter("projectitem2", projectitem2), new SqlParameter("projectitem3", projectitem3), new SqlParameter("projectitem4", projectitem4)).ToList();
            }
            log.Info("get projectitem count=" + lstProject.Count);
            return lstProject;
        }

    }
}

