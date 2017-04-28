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
        PlanService service = new PlanService();

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
            TnderProject service = new TnderProject();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得直接成本資料
            CostAnalysisDataService s = new CostAnalysisDataService();
            List<DirectCost> budget = s.getDirectCost(id);
            ViewBag.result = "共有" + budget.Count + "筆資料"; ;
            return View(budget);
        }
        public String UpdateBudget(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            SYS_USER u = (SYS_USER)Session["user"];
            string msg = "";
            string[] lsttypecode = form.Get("code1").Split(',');
            string[] lsttypesub = form.Get("code2").Split(',');
            string[] lstPrice = form.Get("inputbudget").Split(',');
            List<PLAN_BUDGET> lstItem = new List<PLAN_BUDGET>();
            for (int j = 0; j < lstPrice.Count(); j++)
            {
                PLAN_BUDGET item = new PLAN_BUDGET();
                item.PROJECT_ID = form["id"];
                if (lstPrice[j].ToString() == "")
                {
                    item.BUDGET_AMOUNT = null;
                }
                else
                {
                    item.BUDGET_AMOUNT = decimal.Parse(lstPrice[j]);
                }
                item.TYPE_CODE_1 = lsttypecode[j];
                item.TYPE_CODE_2 = lsttypesub[j];
                item.CREATE_ID = u.USER_ID;
                logger.Debug("Item Project id =" + item.PROJECT_ID + "且九宮格組合為" + item.TYPE_CODE_1 + item.TYPE_CODE_2);
                lstItem.Add(item);
            }
            int i = service.addBudget(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "新增預算資料成功，PROJECT_ID =" + form["id"];
            }

            logger.Info("Request:PROJECT_ID =" + form["id"]);
            return msg;
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
