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
using Newtonsoft.Json;

namespace topmeperp.Controllers
{
    public class EstimationController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        PurchaseFormService service = new PurchaseFormService();


        // GET: Estimation
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName("", "專案執行");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            //畫面上權限管理控制
            //頁面上使用ViewBag 定義開關\@ViewBag.F10005
            //由Session 取得權限清單
            List<SYS_FUNCTION> lstFunctions = (List<SYS_FUNCTION>)Session["functions"];
            //開關預設關閉
            @ViewBag.F10005 = "disabled";
            //輪巡功能清單，若全線存在則將開關打開 @ViewBag.F10005 = "";
            foreach (SYS_FUNCTION f in lstFunctions)
            {
                if (f.FUNCTION_ID == "F10005")
                {
                    @ViewBag.F10005 = "";
                }
            }
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
        //廠商估驗計價
        public ActionResult Valuation(string id)
        {
            logger.Info("valuation index : projectid=" + id);
            ViewBag.projectId = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            List<plansummary> lstContract = null;
            ContractModels contract = new ContractModels();
            lstContract = service.getAllPlanContract(id, Request["supplier"], Request["formname"]);
            contract.contractItems = lstContract;
            return View(contract);
        }

        public ActionResult Search()
        {
            ViewBag.projectId = Request["projectId"];
            TND_PROJECT p = service.getProjectById(Request["projectId"]);
            ViewBag.projectName = p.PROJECT_NAME;
            ContractModels contract = new ContractModels();
            List<plansummary> lstContract = service.getAllPlanContract(Request["projectId"], Request["formname"], Request["supplier"]);
            contract.contractItems = lstContract;
            return View("Valuation", contract);
        }
        
        public ActionResult ContractItems(string id)
        {
            logger.Info("Access To Contract Item By Contract Id =" + id);
            ViewBag.projectId = id.Substring(1, 5).Trim();
            TND_PROJECT p = service.getProjectById(id.Substring(1, 5).Trim());
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.wage = id.Substring(0, 1).Trim();
            ContractModels contract = new ContractModels();
            //取得合約金額與供應商名稱,採購項目等資料
            if (ViewBag.wage == "")
            {
                plansummary lstContract = service.getPlanContract4Est(id.Substring(1).Trim());
                ViewBag.supplier = lstContract.SUPPLIER_ID;
                ViewBag.formname = lstContract.FORM_NAME;
                ViewBag.amount = lstContract.MATERIAL_COST;
                ViewBag.contractid = lstContract.CONTRACT_ID;
                PLAN_PAYMENT_TERMS payment = service.getPaymentTerm(lstContract.CONTRACT_ID);
                if (payment.PAYMENT_RETENTION_RATIO != null)
                {
                    ViewBag.retention = payment.PAYMENT_RETENTION_RATIO ;
                }
                else
                {
                    ViewBag.retention = payment.USANCE_RETENTION_RATIO;
                }
            }
            else
            {
                plansummary lstWageContract = service.getPlanContractOfWage4Est(id.Substring(1).Trim());
                ViewBag.supplier4Wage = lstWageContract.MAN_SUPPLIER_ID;
                ViewBag.formname4Wage = lstWageContract.MAN_FORM_NAME;
                ViewBag.amount4Wage = lstWageContract.WAGE_COST;
                ViewBag.contractid4Wage = lstWageContract.CONTRACT_ID;
                PLAN_PAYMENT_TERMS payment = service.getPaymentTerm(lstWageContract.CONTRACT_ID);
                if (payment.PAYMENT_RETENTION_RATIO != null)
                {
                    ViewBag.retention = payment.PAYMENT_RETENTION_RATIO;
                }
                else
                {
                    ViewBag.retention = payment.USANCE_RETENTION_RATIO;
                }
            }
            ViewBag.date = DateTime.Now;
            List<PLAN_ITEM> lstContractItem = null;
            lstContractItem = service.getContractItemById(id.Substring(1).Trim());
            //contract.planItems = lstContractItem;
            ViewBag.SearchResult = "共取得" + lstContractItem.Count + "筆資料";
            //轉成Json字串
            ViewData["items"] = JsonConvert.SerializeObject(lstContractItem);
            return View("ContractItems", contract);
        }
        //取得合約付款條件
        public string getPaymentTerms(string contractid)
        {
            PurchaseFormService service = new PurchaseFormService();
            logger.Info("access the terms of payment by:" + Request["contractid"]);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getPaymentTerm(contractid));
            logger.Info("plan payment terms info=" + itemJson);
            return itemJson;
        }
    }
    }
