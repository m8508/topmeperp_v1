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
using System.Reflection;

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
            ViewBag.formid = service.getEstNo();
            ViewBag.keyid = id;
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

        //DOM 申購作業功能紐對應不同Action
        public class MultiButtonAttribute : ActionNameSelectorAttribute
        {
            public string Name { get; set; }
            public MultiButtonAttribute(string name)
            {
                this.Name = name;
            }
            public override bool IsValidName(ControllerContext controllerContext,
                string actionName, System.Reflection.MethodInfo methodInfo)
            {
                if (string.IsNullOrEmpty(this.Name))
                {
                    return false;
                }
                return controllerContext.HttpContext.Request.Form.AllKeys.Contains(this.Name);
            }
        }
        [HttpPost]
        [MultiButton("SaveEst")]
        //儲存驗收單(驗收單草稿)
        public ActionResult SaveEst(PLAN_ESTIMATION_FORM est)
        {
            //取得專案編號
            logger.Info("Project Id:" + Request["id"]);
            //取得專案名稱
            logger.Info("Project Name:" + Request["projectName"]);
            logger.Info("ContractId:" + Request["keyid"]);
            //取得合約估驗品項ID
            string[] lstItemId = Request["planitemid"].Split(',');
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                logger.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] lstQty = Request["evaluated_qty"].Split(',');
            //建立估驗單
            logger.Info("create new Estimation Form");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            est.PROJECT_ID = Request["id"];
            est.CREATE_ID = u.USER_ID;
            est.CREATE_DATE = DateTime.Now;
            est.EST_FORM_ID = Request["formid"];
            est.CONTRACT_ID = Request["contractid"];
            est.PLUS_TAX = Request["tax"];
            est.TAX_AMOUNT = int.Parse(Request["taxAmount"]);
            est.STATUS = 0;
            est.PAYMENT_TRANSFER = int.Parse(Request["totalAmount"]);
            est.FOREIGN_PAYMENT = int.Parse(Request["foreign_payment"]);
            est.DEDUCTED_ADVANCE_PAYMENT = int.Parse(Request["advanceAmount"]);
            est.REMARK = Request["remark"];
            est.RETENTION_PAYMENT = int.Parse(Request["retentionAmount"]);
            PLAN_ESTIMATION_FORM item = new PLAN_ESTIMATION_FORM();
            string estid = service.newEST(Request["formid"], est, lstItemId);
            List<PLAN_ESTIMATION_ITEM> lstItem = new List<PLAN_ESTIMATION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_ESTIMATION_ITEM items = new PLAN_ESTIMATION_ITEM();
                items.PLAN_ITEM_ID = lstItemId[j];
                if (lstQty[j].ToString() == "")
                {
                    items.EST_QTY = null;
                }
                else
                {
                    items.EST_QTY = decimal.Parse(lstQty[j]);
                }
                logger.Debug("Item No=" + items.PLAN_ITEM_ID + ", Qty =" + items.EST_QTY);
                lstItem.Add(items);
            }
            int k = service.refreshEST(estid, est, lstItem);
            return Redirect("ContractItems?id=" + Request["keyid"]);
        }

        //估驗單查詢
        public ActionResult EstimationForm(string id)
        {
            logger.Info("Search For Estimation Form !!");
            ViewBag.projectid = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //估驗單草稿
            int status = 10;
            if (Request["status"] == null || Request["status"] == "")
            {
                status = 0;
            }
            List<ESTFunction> lstEST = service.getESTByEstId(id, Request["contractid"], Request["estid"], status);
            return View(lstEST);
        }

        public ActionResult SearchEST()
        {
            logger.Info("projectid=" + Request["id"] + ", contractid =" + Request["contractid"] + ", estid =" + Request["estid"] + ", status =" + int.Parse(Request["status"]));
            List<ESTFunction> lstEST = service.getESTByEstId(Request["id"], Request["contractid"], Request["estid"], int.Parse(Request["status"]));
            ViewBag.SearchResult = "共取得" + lstEST.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("EstimationForm", lstEST);
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

        //顯示單一估驗單功能
        public ActionResult SingleEST(string id)
        {
            logger.Info("http get mehtod:" + id);
            EstimationFormDetail singleForm = new EstimationFormDetail();
            service.getPRByPrId(id);
            singleForm.planEST = service.formEST;
            singleForm.planESTItem = service.ESTItem;
            singleForm.prj = service.getProjectById(singleForm.planEST.PROJECT_ID);
            logger.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            return View(singleForm);
        }
    }
    }
