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
using System.Data;
using System.Globalization;

namespace topmeperp.Controllers
{
    public class MaterialManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(InquiryController));
        PurchaseFormService service = new PurchaseFormService();
        ProjectPlanService planService = new ProjectPlanService();

        // GET: MaterialManage
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
                log.Info("search project by 名稱 =" + projectname);
                List<topmeperp.Models.TND_PROJECT> lstProject = new List<TND_PROJECT>();
                using (var context = new topmepEntities())
                {
                    lstProject = context.TND_PROJECT.SqlQuery("select * from TND_PROJECT p "
                        + "where p.PROJECT_NAME Like '%' + @projectname + '%' AND STATUS=@status;",
                         new SqlParameter("projectname", projectname), new SqlParameter("status", status)).ToList();
                }
                log.Info("get project count=" + lstProject.Count);
                return lstProject;
            }
            else
            {
                return null;
            }
        }

        //Plan Task 任務連結物料
        public ActionResult PlanTask(string id)
        {
            log.Debug("show sreen for apply for material");
            ViewBag.projectid = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.TreeString = planService.getProjectTask4Tree(id); ;
            return View();
        }

        //物料申購
        public ActionResult Application(string id)
        {
            log.Info("Access to Application page!!");
            ViewBag.projectid = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            return View();
        }

        //申購單查詢
        public ActionResult PurchaseRequisition(string id)
        {
            log.Info("Access to Purchase Requisition page!!");
            ViewBag.projectid = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            //取得材料合約供應商資料
            List<SelectListItem> selectMain = new List<SelectListItem>();
            foreach (string itm in service.getFormNameForContract(id))
            {
                log.Debug("supplier=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectMain.Add(selectI);
                }
            }
            // selectMain.Add(empty);
            ViewBag.formname = selectMain;
            return View();
        }
        
        [HttpPost]
        public ActionResult PurchaseRequisition(FormCollection f)
        {

            log.Info("projectid=" + Request["id"]);
            List<topmeperp.Models.PLAN_ITEM> lstProject = service.getPlanItem(Request["id"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formName"], Request["supplier"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            //取得材料合約供應商資料
            List<SelectListItem> selectMain = new List<SelectListItem>();
            foreach (string itm in service.getFormNameForContract(Request["id"]))
            {
                log.Debug("supplier=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectMain.Add(selectI);
                }
            }
            // selectMain.Add(empty);
            ViewBag.formname = selectMain;
            return View("PurchaseRequisition", lstProject);
        }
    }
}
