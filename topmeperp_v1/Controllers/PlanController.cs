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
    public class PlanController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
     

        // GET: Plan
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName("", "專案執行");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View(lstProject);
        }

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
        public ActionResult Budget(string id)
        {
            logger.Info("budget info for projectid=" + id);
            ViewBag.projectid = id;
            //取得直接成本資料
            CostAnalysisDataService service = new CostAnalysisDataService();
            List<DirectCost> budget = service.getDirectCost(id);
            return View(budget);
        }

        // GET: Purchasing/Details/5
        public ActionResult PurchaseMain(string id)
        {
            //傳入專案編號，
            logger.Info("start project id=" + id);

            //取得專案基本資料
            PurchaseService service = new PurchaseService();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.id = p.PROJECT_ID;
            ViewBag.projectName = p.PROJECT_NAME;
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            //取得主系統資料
            List<SelectListItem> selectMain = new List<SelectListItem>();
            foreach (string itm in service.getSystemMain(id))
            {
                logger.Debug("Main System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectMain.Add(selectI);
                }
            }
            // selectMain.Add(empty);
            ViewBag.SystemMain = selectMain;
            //取得次系統資料
            List<SelectListItem> selectSub = new List<SelectListItem>();
            foreach (string itm in service.getSystemSub(id))
            {
                logger.Debug("Sub System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectSub.Add(selectI);
                }
            }
            //selectSub.Add(empty);
            ViewBag.SystemSub = selectSub;
            //設定查詢條件
            return View();
        }
    
    }
}
