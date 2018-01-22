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
            List<ProjectList> lstProject = SearchProjectByName("", "專案執行");
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

        private List<ProjectList> SearchProjectByName(string projectname, string status)
        {
            if (projectname != null)
            {
                logger.Info("search project by 名稱 =" + projectname);
                List<ProjectList> lstProject = new List<ProjectList>();
                using (var context = new topmepEntities())
                {
                    lstProject = context.Database.SqlQuery<ProjectList>("select DISTINCT p.*, convert(varchar, pi.CREATE_DATE , 111) as PLAN_CREATE_DATE from TND_PROJECT p left join PLAN_ITEM pi "
                        + "on p.PROJECT_ID = pi.PROJECT_ID where p.PROJECT_NAME Like '%' + @projectname + '%' AND STATUS=@status;",
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

        public ActionResult AddEST(string id, string projectid, string type)
        {
            logger.Info("create EST form id process!  Contract id=" + id);
            //新增估驗單編號
            string formid = service.getEstNo();
            return RedirectToAction("ContractItems", "Estimation", new { id = id, formid = formid, projectid = projectid, type = type });
        }

        public ActionResult ContractItems(string id, string formid, string projectid, string type)
        {
            logger.Info("Access To Contract Item By Contract Id =" + id);
            ViewBag.projectId = projectid;
            TND_PROJECT p = service.getProjectById(projectid);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.wage = type;
            ContractModels contract = new ContractModels();
            ViewBag.formid = formid;
            ViewBag.contractid = id;
            ViewBag.keyid = id; //使用供應商名稱的contractid
            //取得合約金額與供應商名稱,採購項目等資料
            if (ViewBag.wage != "Y")
            {
                plansummary lstContract = service.getPlanContract4Est(id);
                ViewBag.supplier = lstContract.SUPPLIER_ID;
                ViewBag.formname = lstContract.FORM_NAME;
                ViewBag.amount = lstContract.MATERIAL_COST;
                //ViewBag.contractid = lstContract.CONTRACT_ID; //使用供應商編號的contractid
                PLAN_PAYMENT_TERMS payment = service.getPaymentTerm(id);
                if (payment.PAYMENT_RETENTION_RATIO != null)
                {
                    ViewBag.retention = payment.PAYMENT_RETENTION_RATIO;
                }
                else
                {
                    ViewBag.retention = payment.USANCE_RETENTION_RATIO;
                }
                if (payment.PAYMENT_TYPE == "連工帶料")
                {
                    ViewBag.type = 'C'; // 合約包含材料與工資
                }
                else
                {
                    ViewBag.type = 'M'; // 設備合約
                }
                ViewBag.estCount = service.getEstCountById(id);
                ViewBag.paymentTerms = service.getTermsByContractId(id);
                var balance = service.getBalanceOfRefundById(id);
                if (balance > 0)
                {
                    TempData["balance"] = "本合約目前尚有 " + string.Format("{0:C0}", balance) + "的代付支出款項，仍未扣回!";
                }
            }
            else
            {
                plansummary lstWageContract = service.getPlanContractOfWage4Est(id);
                ViewBag.supplier4Wage = lstWageContract.MAN_SUPPLIER_ID;
                ViewBag.formname4Wage = lstWageContract.MAN_FORM_NAME;
                ViewBag.amount4Wage = lstWageContract.WAGE_COST;
                ViewBag.contractid4Wage = lstWageContract.CONTRACT_ID;
                ViewBag.type = 'W'; //工資合約
                PLAN_PAYMENT_TERMS payment = service.getPaymentTerm(lstWageContract.CONTRACT_ID);
                if (payment.PAYMENT_RETENTION_RATIO != null)
                {
                    ViewBag.retention = payment.PAYMENT_RETENTION_RATIO;
                }
                else
                {
                    ViewBag.retention = payment.USANCE_RETENTION_RATIO;
                }
                ViewBag.estCount = service.getEstCountById(lstWageContract.CONTRACT_ID);
                ViewBag.paymentTerms = service.getTermsByContractId(lstWageContract.CONTRACT_ID);
                var balance = service.getBalanceOfRefundById(lstWageContract.CONTRACT_ID);
                if (balance > 0)
                {
                    TempData["balance"] = "本合約目前尚有 " + string.Format("{0:C0}", balance) + "的代付支出款項，仍未扣回!";
                }
            }
            ViewBag.date = DateTime.Now;
            //ViewBag.paymentkey = ViewBag.formid + ViewBag.contractid;
            List<EstimationForm> lstContractItem = null;
            lstContractItem = service.getContractItemById(id, projectid);
            //contract.planItems = lstContractItem;
            ViewBag.SearchResult = "共取得" + lstContractItem.Count + "筆資料";
            //轉成Json字串
            ViewData["items"] = JsonConvert.SerializeObject(lstContractItem);
            return View("ContractItems", contract);
        }

        //DOM 估驗作業功能紐對應不同Action
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
        public ActionResult SaveEst(FormCollection f)
        {
            //取得估驗單編號
            logger.Info("EST form Id:" + Request["formid"]);
            logger.Info("Project Id:" + Request["projectid"]);
            logger.Info("get EST No of un Approval By contractid:" + Request["contractid"]);
            string contractid = Request["contractid"];
            string UnApproval = null;
            UnApproval = service.getEstNoByContractId(contractid);
            if (UnApproval != null && "" != UnApproval)
            {
                TempData["result"] = "目前尚有未核准的估驗單，估驗單編號為" + UnApproval + "，待此單核准後再新增估驗單!";
                return RedirectToAction("Valuation", "Estimation", new { id = Request["projectid"] });
            }
            else
            {
                //更新估驗單
                logger.Info("update Estimation Form");
                //估驗單草稿(未送審) STATUS = 10
                int k = service.UpdateESTStatusById(Request["formid"]);
                return RedirectToAction("SingleEST", "Estimation", new { id = Request["formid"] });
            }
        }

        [HttpPost]
        [MultiButton("AddEst")]
        //驗收單送審
        public ActionResult AddEst(FormCollection f)
        {
            //取得估驗單編號
            logger.Info("EST form Id:" + Request["formid"]);
            logger.Info("get EST No of un Approval By contractid:" + Request["contractid"]);
            logger.Info("Project Id:" + Request["projectid"]);
            string contractid = Request["contractid"];
            string UnApproval = null;
            UnApproval = service.getEstNoByContractId(contractid);
            if (UnApproval != null && "" != UnApproval)
            {
                TempData["result"] = "目前尚有未核准的估驗單，估驗單編號為" + UnApproval + "，待此單核准後再新增估驗單!";
                return RedirectToAction("Valuation", "Estimation", new { id = Request["projectid"] });
            }
            else
            {
                //更新估驗單
                logger.Info("update Estimation Form");
                //估驗單(已送審) STATUS = 20
                int k = service.RefreshESTStatusById(Request["formid"]);
                return RedirectToAction("SingleEST", "Estimation", new { id = Request["formid"] });
            }
        }

        public String ConfirmEst(PLAN_ESTIMATION_FORM est)
        {
            //取得專案編號
            logger.Info("Project Id:" + Request["id"]);
            //取得專案名稱
            logger.Info("Project Name:" + Request["projectName"]);
            //logger.Info("ContractId:" + Request["keyid"]);
            logger.Info("get EST No of un Approval By contractid:" + Request["contractid"]);
            string contractid = Request["contractid"];
            string UnApproval = null;
            UnApproval = service.getEstNoByContractId(contractid);
            if (UnApproval != null && "" != UnApproval)
            {
                System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                string itemJson = objSerializer.Serialize(service.getEstNoByContractId(contractid));
                logger.Info("EST form No of UnApproval =" + itemJson);
                return itemJson;
            }
            else
            {
                // 先刪除原先資料
                logger.Info("EST form id =" + Request["formid"]);
                logger.Info("Delete PLAN_ESTIMATION_FORM By EST_FORM_ID");
                service.delESTByESTId(Request["formid"]);
                logger.Info("Delete PLAN_ESTIMATION_ITEM By EST_FORM_ID");
                service.delESTItemsByESTId(Request["formid"]);
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
                est.REMARK = Request["remark"];
                est.INVOICE = Request["invoice"];
                if (Request["tax"] == "E")
                {
                    est.TAX_RATIO = decimal.Parse(Request["taxratio"]);
                }
                else
                {
                    est.TAX_RATIO = 0;
                }
                try
                {
                    est.FOREIGN_PAYMENT = int.Parse(Request["t_foreign"]);
                }
                catch (Exception ex)
                {
                    logger.Error(est.FOREIGN_PAYMENT + " not foreign_payment:" + ex.Message);
                }
                est.TYPE = Request["type"];
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
                int m = service.UpdateRetentionAmountById(Request["formid"], Request["contractid"]);
                System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                string itemJson = objSerializer.Serialize(service.getDetailsPayById(Request["formid"], Request["contractid"]));
                logger.Info("EST form details payment amount=" + itemJson);
                return itemJson;
            }
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
            int status = 20;
            if (Request["status"] == null || Request["status"] == "")
            {
                status = 10;
            }
            List<ESTFunction> lstEST = service.getESTListByEstId(id, Request["contractid"], Request["estid"], status);
            return View(lstEST);
        }

        public ActionResult SearchEST()
        {
            logger.Info("projectid=" + Request["id"] + ", contractid =" + Request["contractid"] + ", estid =" + Request["estid"] + ", status =" + int.Parse(Request["status"]));
            List<ESTFunction> lstEST = service.getESTListByEstId(Request["id"], Request["contractid"], Request["estid"], int.Parse(Request["status"]));
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
            ContractModels singleForm = new ContractModels();
            service.getESTByEstId(id);
            singleForm.planEST = service.formEST;
            ViewBag.wage = singleForm.planEST.TYPE;
            ViewBag.contractid = singleForm.planEST.CONTRACT_ID;
            TND_SUPPLIER s = service.getSupplierInfo(singleForm.planEST.CONTRACT_ID.Substring(6, 7).Trim());
            ViewBag.supplier = s.COMPANY_NAME;
            if (ViewBag.wage != "W")
            {
                plansummary lstContract = service.getPlanContract4Est(singleForm.planEST.CONTRACT_ID.Replace(singleForm.planEST.CONTRACT_ID.Substring(6, 7), s.COMPANY_NAME));
                ViewBag.contractamount = lstContract.MATERIAL_COST;
            }
            else
            {
                plansummary lstWageContract = service.getPlanContractOfWage4Est(singleForm.planEST.CONTRACT_ID.Replace(singleForm.planEST.CONTRACT_ID.Substring(6, 7), s.COMPANY_NAME));
                ViewBag.contractamount = lstWageContract.WAGE_COST;
            }
            PLAN_PAYMENT_TERMS payment = service.getPaymentTerm(singleForm.planEST.CONTRACT_ID);
            if (payment.PAYMENT_RETENTION_RATIO != null)
            {
                ViewBag.retention = payment.PAYMENT_RETENTION_RATIO;
            }
            else
            {
                ViewBag.retention = payment.USANCE_RETENTION_RATIO;
            }
            ViewBag.paymentTerms = service.getTermsByContractId(singleForm.planEST.CONTRACT_ID);
            ViewBag.estCount = service.getEstCountByESTId(id);
            ViewBag.formname = singleForm.planEST.CONTRACT_ID.Substring(13).Trim();
            ViewBag.paymentkey = id + singleForm.planEST.CONTRACT_ID;
            singleForm.planESTItem = service.ESTItem;
            singleForm.prj = service.getProjectById(singleForm.planEST.PROJECT_ID);
            logger.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            PaymentDetailsFunction lstSummary = service.getDetailsPayById(id, singleForm.planEST.CONTRACT_ID);
            var balance = service.getBalanceOfRefundById(ViewBag.contractid);
            if (balance > 0)
            {
                TempData["balance"] = "本合約目前尚有 " + string.Format("{0:C0}", balance) + "的代付支出款項，仍未扣回!";
            }
            //轉成Json字串
            ViewData["items"] = JsonConvert.SerializeObject(singleForm.planESTItem);
            ViewData["summary"] = JsonConvert.SerializeObject(lstSummary);
            return View(singleForm);
        }
        //其他扣款
        public ActionResult OtherPayment(string id, string contractid)
        {
            logger.Info("Access To Other Payment By EST Form Id =" + id);
            service.getInqueryForm(id);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.projectId = f.PROJECT_ID;
            TND_PROJECT p = service.getProjectById(f.PROJECT_ID);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = contractid;
            ViewBag.formname = f.FORM_NAME;
            ViewBag.formid = id;
            List<PLAN_OTHER_PAYMENT> lstOtherPayItem = null;
            lstOtherPayItem = service.getOtherPayById(id);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstOtherPayItem.Count;
            logger.Debug("this other payment record =" + ViewBag.key + "筆");
            ViewData["items"] = JsonConvert.SerializeObject(lstOtherPayItem);
            return View();
        }
        public String AddOtherPay(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 取得其他扣款資料
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstReason = form.Get("input_reason").Split(',');
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Other Payment Amount  =" + item.AMOUNT);
                item.REASON = lstReason[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且扣款原因為" + item.REASON);
                lstItem.Add(item);
            }
            int i = service.addOtherPayment(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "新增其他扣款資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }

        public String UpdateOtherPay(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 先刪除原先資料
            logger.Info("EST form id =" + form["formid"]);
            logger.Info("Delete PLAN_OTHER_PAYMENT By EST_FORM_ID");
            service.delOtherPayByESTId(form["formid"]);
            // 再次取得其他扣款資料
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstReason = form.Get("input_reason").Split(',');
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Other Payment Amount  =" + item.AMOUNT);
                item.REASON = lstReason[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且扣款原因為" + item.REASON);
                lstItem.Add(item);
            }
            int i = service.addOtherPayment(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新其他扣款資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }

        //更新估驗數量與稅率
        public String UpdateESTQty(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            string tax = form.Get("tax").Trim();
            decimal taxratio = 0;
            if (tax == "E")
            {
                taxratio = decimal.Parse(form.Get("taxratio").Trim());
            }
            string[] lstPlanItemId = form.Get("planitemid").Split(',');
            string[] lstQty = form.Get("evaluated_qty").Split(',');
            List<PLAN_ESTIMATION_ITEM> lstItem = new List<PLAN_ESTIMATION_ITEM>();
            for (int j = 0; j < lstPlanItemId.Count(); j++)
            {
                PLAN_ESTIMATION_ITEM item = new PLAN_ESTIMATION_ITEM();
                item.PLAN_ITEM_ID = lstPlanItemId[j];
                if (lstQty[j].ToString() == "")
                {
                    item.EST_QTY = null;
                }
                else
                {
                    item.EST_QTY = decimal.Parse(lstQty[j]);
                }
                logger.Debug("EST_FIRM_ID = " + form["estid"] + "It's Plan tem Id No=" + item.PLAN_ITEM_ID + ", EST Qty =" + item.EST_QTY);
                lstItem.Add(item);
            }
            int k = service.RefreshESTByEstId(form["estid"], tax, taxratio);
            int i = service.refreshESTQty(form["estid"], lstItem);
            int m = service.UpdateRetentionAmountById(form["estid"], form["contractid"]);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新估驗數量成功，請確認各項金額是否正確 !";
            }

            logger.Info("Request: 更新數量訊息 = " + msg);
            return msg;
        }

        //更新估驗單
        public String UpdateEST(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            decimal foreign_payment = decimal.Parse(form.Get("t_foreign").Trim());
            decimal retention = decimal.Parse(form.Get("t_retention").Trim());
            decimal tax_amount = decimal.Parse(form.Get("tax_amount").Trim());
            string remark = form.Get("remark").Trim();
            int i = service.RefreshESTAmountByEstId(form["estid"], foreign_payment, retention, tax_amount, remark);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新估驗單成功";
            }

            logger.Info("Request: 更新估驗單訊息 = " + msg);
            return msg;
        }

        //預付款
        public ActionResult AdvancePayment(string id, string contractid)
        {
            logger.Info("Access To Advance Payment By EST Form Id =" + id);
            service.getInqueryForm(id);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.projectId = f.PROJECT_ID;
            TND_PROJECT p = service.getProjectById(f.PROJECT_ID);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = contractid;
            ViewBag.formid = id;
            PLAN_PAYMENT_TERMS payment = service.getPaymentTerm(contractid);
            ViewBag.advancePaymentRatio = payment.PAYMENT_ADVANCE_RATIO;
            AdvancePaymentFunction advancePay = service.getAdvancePayById(id, contractid);
            List<PLAN_OTHER_PAYMENT> lstAdvancePayItem = null;
            lstAdvancePayItem = service.getAdvancePayByESTId(id);
            ViewBag.key = lstAdvancePayItem.Count;
            ViewData["items"] = JsonConvert.SerializeObject(advancePay);
            return View();
        }

        public String AddAdvancePay(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            string advance_payment = form.Get("advance_payment");
            string temporary_loan = form.Get("temporary_loan");
            string margins = form.Get("margins");
            string[] lstItemId = String.Join(",", advance_payment, temporary_loan, margins).Split(',');
            string[] lsttype = String.Join(",", "A", "B", "C").Split(',');
            // 取得預付款資料
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                if (lstItemId[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstItemId[j]);
                }
                logger.Info("Advance Payment Amount  =" + item.AMOUNT);
                item.TYPE = lsttype[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + " ,And Type =" + item.TYPE);
                lstItem.Add(item);
            }
            int i = service.addAdvancePayment(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "新增預付款資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }

        public String UpdateAdvancePay(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 取得預付款資料
            string advance_payment = form.Get("advance_payment");
            string temporary_loan = form.Get("temporary_loan");
            string margins = form.Get("margins");
            string[] lstItemId = String.Join(",", advance_payment, temporary_loan, margins).Split(',');
            string[] lsttype = String.Join(",", "A", "B", "C").Split(',');
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                if (lstItemId[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstItemId[j]);
                }
                logger.Info("Advance Payment Amount  =" + item.AMOUNT);
                item.TYPE = lsttype[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + " ,And Type =" + item.TYPE);
                lstItem.Add(item);
            }
            int i = service.updateAdvancePayment(form["formid"], lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "修改預付款資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }

        public String UpdateESTStatusById(FormCollection form)
        {
            //取得估驗單編號
            logger.Info("form:" + form.Count);
            logger.Info("EST form Id:" + form["estid"]);
            string msg = "";
            decimal foreign_payment = decimal.Parse(form.Get("t_foreign").Trim());
            decimal retention = decimal.Parse(form.Get("t_retention").Trim());
            decimal tax_amount = decimal.Parse(form.Get("tax_amount").Trim());
            string remark = form.Get("remark").Trim();
            int i = service.RefreshESTAmountByEstId(form["estid"], foreign_payment, retention, tax_amount, remark);
            //更新估驗單狀態
            logger.Info("Update Estimation Form Status");
            //估驗單(已送審) STATUS = 20
            int j = service.RefreshESTStatusById(form["estid"]);
            if (j == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "估驗單已送審";
            }
            return msg;
        }

        public String RejectESTById(FormCollection form)
        {
            //取得估驗單編號
            logger.Info("EST form Id:" + form["estid"]);
            //更新估驗單狀態
            logger.Info("Reject Estimation Form ");
            //估驗單(已退件) STATUS = 0
            string msg = "";
            int i = service.RejectESTByEstId(form["estid"]);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "估驗單已退回";
            }
            return msg;
        }

        public String ApproveESTById(FormCollection form)
        {
            //取得估驗單編號
            logger.Info("EST form Id:" + form["estid"]);
            //更新估驗單狀態
            logger.Info("Confirm Estimation Form ");
            //估驗單(已核可) STATUS = 30
            string msg = "";
            int i = service.ApproveESTByEstId(form["estid"]);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "估驗單已核可";
            }
            return msg;
        }

        //憑證
        public ActionResult Invoice(string id, string contractid)
        {
            logger.Info("Access To Invoice By EST Form Id =" + id);
            service.getInqueryForm(id);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.projectId = f.PROJECT_ID;
            TND_PROJECT p = service.getProjectById(f.PROJECT_ID);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.formname = f.FORM_NAME;
            ViewBag.contractid = contractid;
            ViewBag.formid = id;
            TND_SUPPLIER lstSupplier = service.getSupplierInfo(f.SUPPLIER_ID);
            ViewBag.supplier = lstSupplier.COMPANY_NAME;
            ViewBag.companyid = lstSupplier.COMPANY_ID;
            ContractModels singleForm = new ContractModels();
            service.getESTByEstId(id);
            singleForm.planEST = service.formEST;
            try
            {
                ViewBag.type = "不檢查發票";
                if (singleForm.planEST.INVOICE == "E")
                {
                    ViewBag.type = "發票含保留款";
                }
                else if (singleForm.planEST.INVOICE == "I")
                {
                    ViewBag.type = "發票不含保留款";
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
            }
            try
            {
                ViewBag.tax = "其他";
                if (singleForm.planEST.PLUS_TAX == "E")
                {
                    ViewBag.tax = "外加稅 " + singleForm.planEST.TAX_RATIO + "  %";
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
            }
            List<PLAN_INVOICE> lstInvoice = null;
            lstInvoice = service.getInvoiceById(id);
            ViewBag.key = lstInvoice.Count;
            logger.Debug("this invoice record =" + ViewBag.key + "筆");
            ViewData["items"] = JsonConvert.SerializeObject(lstInvoice);
            return View();
        }

        public String AddInvoice(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 取得憑證資料
            string[] lstDate = form.Get("invoice_date").Split(',');
            string[] lstNumber = form.Get("invoice_number").Split(',');
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstTax = form.Get("taxamount").Split(',');
            string[] lstType = form.Get("invoicetype").Split(',');
            string[] lstSubType = form.Get("sub_type").Split(',');
            string[] lstPlanItem = form.Get("plan_item_id").Split(',');
            string[] lstDiscountQty = form.Get("discount_qty").Split(',');
            string[] lstDiscountPrice = form.Get("discount_unit_price").Split(',');
            List<PLAN_INVOICE> lstItem = new List<PLAN_INVOICE>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_INVOICE item = new PLAN_INVOICE();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                item.INVOICE_NUMBER = lstNumber[j];
                item.INVOICE_DATE = Convert.ToDateTime(lstDate[j]);
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                if (lstTax[j].ToString() == "")
                {
                    item.TAX = null;
                }
                else
                {
                    item.TAX = decimal.Parse(lstTax[j]);
                }
                item.TYPE = lstType[j];
                item.SUB_TYPE = lstSubType[j];
                item.PLAN_ITEM_ID = lstPlanItem[j];
                if (lstDiscountQty[j].ToString() == "")
                {
                    item.DISCOUNT_QTY = null;
                }
                else
                {
                    item.DISCOUNT_QTY = decimal.Parse(lstDiscountQty[j]);
                }
                if (lstDiscountPrice[j].ToString() == "")
                {
                    item.DISCOUNT_UNIT_PRICE = null;
                }
                else
                {
                    item.DISCOUNT_UNIT_PRICE = decimal.Parse(lstDiscountPrice[j]);
                }
                logger.Info("Invoice Number = " + item.INVOICE_NUMBER + "and Invoice Amount =" + item.AMOUNT);
                //logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且憑證類型為" + item.TYPE);
                lstItem.Add(item);
            }
            int i = service.addInvoice(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "新增憑證資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }

        public String UpdateInvoice(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 先刪除原先資料
            logger.Info("EST form id =" + form["formid"]);
            logger.Info("Delete PLAN_INVOICE By EST_FORM_ID");
            service.delInvoiceByESTId(form["formid"]);
            // 再次取得憑證資料
            string[] lstDate = form.Get("invoice_date").Split(',');
            string[] lstNumber = form.Get("invoice_number").Split(',');
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstTax = form.Get("taxamount").Split(',');
            string[] lstType = form.Get("invoicetype").Split(',');
            string[] lstSubType = form.Get("sub_type").Split(',');
            string[] lstPlanItem = form.Get("plan_item_id").Split(',');
            string[] lstDiscountQty = form.Get("discount_qty").Split(',');
            string[] lstDiscountPrice = form.Get("discount_unit_price").Split(',');
            List<PLAN_INVOICE> lstItem = new List<PLAN_INVOICE>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_INVOICE item = new PLAN_INVOICE();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                item.INVOICE_NUMBER = lstNumber[j];
                item.INVOICE_DATE = Convert.ToDateTime(lstDate[j]);
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                if (lstTax[j].ToString() == "")
                {
                    item.TAX = null;
                }
                else
                {
                    item.TAX = decimal.Parse(lstTax[j]);
                }
                item.TYPE = lstType[j];
                if (lstType[j] == "折讓單")
                {
                    item.SUB_TYPE = lstSubType[j];
                    item.PLAN_ITEM_ID = lstPlanItem[j];
                    if (lstDiscountQty[j].ToString() == "")
                    {
                        item.DISCOUNT_QTY = null;
                    }
                    else
                    {
                        item.DISCOUNT_QTY = decimal.Parse(lstDiscountQty[j]);
                    }
                    if (lstDiscountPrice[j].ToString() == "")
                    {
                        item.DISCOUNT_UNIT_PRICE = null;
                    }
                    else
                    {
                        item.DISCOUNT_UNIT_PRICE = decimal.Parse(lstDiscountPrice[j]);
                    }
                }
                else
                {
                    item.SUB_TYPE = null;
                    item.PLAN_ITEM_ID = null;
                    item.DISCOUNT_QTY = null;
                    item.DISCOUNT_UNIT_PRICE = null;
                }
                logger.Info("Invoice Number = " + item.INVOICE_NUMBER + "and Invoice Amount =" + item.AMOUNT);
                lstItem.Add(item);
            }
            int i = service.addInvoice(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新憑證資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }

        //代付支出
        public ActionResult RePayment(string id, string contractid)
        {
            logger.Info("Access To RePayment By EST Form Id =" + id);
            service.getInqueryForm(id);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.projectId = f.PROJECT_ID;
            TND_PROJECT p = service.getProjectById(f.PROJECT_ID);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = contractid;
            ViewBag.formname = f.FORM_NAME;
            ViewBag.formid = id;
            List<RePaymentFunction> lstOtherPayItem = null;
            lstOtherPayItem = service.getRePaymentById(id);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstOtherPayItem.Count;
            logger.Debug("this repayment record =" + ViewBag.key + "筆");
            ViewData["items"] = JsonConvert.SerializeObject(lstOtherPayItem);
            logger.Debug(ViewData["items"]);
            return View();
        }
        //代付支出-選商
        public ActionResult ChooseSupplier(string id, string contractid)
        {
            logger.Info("Access To RePayment By EST Form Id =" + id);
            service.getInqueryForm(id);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.contractid = contractid;
            ViewBag.formid = id;
            List<RePaymentFunction> lstSupplier = null;
            lstSupplier = service.getSupplierOfContractByPrjId(f.PROJECT_ID);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstSupplier.Count;
            logger.Debug("supplier record =" + ViewBag.key + "筆");
            return View(lstSupplier);
        }

        public ActionResult AddRePayment(FormCollection form)
        {
            //取得估驗單編號
            logger.Info("EST Form Id:" + Request["formid"]);
            //取得合約編號
            logger.Info("Contract Id:" + Request["contractid"]);
            //取得使用者勾選品項ID
            logger.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            logger.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                logger.Info("item_list return No.:" + lstItemId[i]);
            }
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = Request["formid"];
                item.CONTRACT_ID = Request["contractid"];
                item.CONTRACT_ID_FOR_REFUND = lstItemId[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + " ,and supplier for refund =" + item.CONTRACT_ID_FOR_REFUND);
                lstItem.Add(item);
            }
            int k = service.AddRePay(lstItem);
            return Redirect("RePayment?id=" + Request["formid"] + "&contractid=" + Request["contractid"]);
        }
        public String UpdateRePay(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 先刪除原先資料
            logger.Info("EST form id =" + form["formid"]);
            logger.Info("Delete PLAN_OTHER_PAYMENT By EST_FORM_ID");
            service.delRePayByESTId(form["formid"]);
            // 再次取得代付支出資料
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstReason = form.Get("input_reason").Split(',');
            string[] lstContract4Refund = form.Get("contractid4refund").Split(',');
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Other Payment Amount  =" + item.AMOUNT);
                item.REASON = lstReason[j];
                item.CONTRACT_ID_FOR_REFUND = lstContract4Refund[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且扣款原因為" + item.REASON + "且代付支出對象為" + item.CONTRACT_ID_FOR_REFUND);
                lstItem.Add(item);
            }
            int i = service.AddRePay(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新代付支出資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }
        //代付扣回
        public ActionResult Refund(string id, string contractid)
        {
            logger.Info("Access To Refund By EST Form Id =" + id);
            service.getInqueryForm(id);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.projectId = f.PROJECT_ID;
            TND_PROJECT p = service.getProjectById(f.PROJECT_ID);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = contractid;
            ViewBag.formname = f.FORM_NAME;
            ViewBag.formid = id;
            List<RePaymentFunction> lstOtherPayItem = null;
            List<RePaymentFunction> lstRefundItem = null;
            lstOtherPayItem = service.getRefundById(id);
            lstRefundItem = service.getRefundOfSupplierById(contractid);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstOtherPayItem.Count;
            logger.Debug("this repayment record =" + ViewBag.key + "筆");
            ViewData["items"] = JsonConvert.SerializeObject(lstOtherPayItem);
            logger.Debug(ViewData["items"]);
            return View(lstRefundItem);
        }

        //代付支出-選商
        public ActionResult ChooseSupplierOfRefund(string id, string contractid)
        {
            logger.Info("Access To Refund By EST Form Id =" + id);
            service.getInqueryForm(id);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.contractid = contractid;
            ViewBag.formid = id;
            List<RePaymentFunction> lstSupplier = null;
            lstSupplier = service.getSupplierOfContractRefundById(contractid);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstSupplier.Count;
            logger.Debug("supplier record =" + ViewBag.key + "筆");
            return View(lstSupplier);
        }

        public ActionResult AddRefund(FormCollection form)
        {
            //取得估驗單編號
            logger.Info("EST Form Id:" + Request["formid"]);
            //取得合約編號
            logger.Info("Contract Id:" + Request["contractid"]);
            //取得使用者勾選品項ID
            logger.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            logger.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                logger.Info("item_list return No.:" + lstItemId[i]);
            }
            PLAN_OTHER_PAYMENT Item = new PLAN_OTHER_PAYMENT();
            string k = service.AddRefund(Request["formid"], Request["contractid"], lstItemId);
            return Redirect("Refund?id=" + Request["formid"] + "&contractid=" + Request["contractid"]);
        }

        public String UpdateRefund(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 先刪除原先資料
            logger.Info("EST form id =" + form["formid"]);
            logger.Info("Delete PLAN_OTHER_PAYMENT By EST_FORM_ID");
            service.delRefundByESTId(form["formid"]);
            // 再次取得代付扣回資料
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstReason = form.Get("input_reason").Split(',');
            string[] lstEstCount = form.Get("est_count").Split(',');
            string[] lstEstIdRefund = form.Get("est_id_refund").Split(',');
            string[] lstContract4Refund = form.Get("contractid4refund").Split(',');
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                if (lstEstCount[j].ToString() == "")
                {
                    item.EST_COUNT_REFUND = null;
                }
                else
                {
                    item.EST_COUNT_REFUND = int.Parse(lstEstCount[j]);
                }
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Other Payment Amount  =" + item.AMOUNT);
                item.REASON = lstReason[j];
                item.CONTRACT_ID_FOR_REFUND = lstContract4Refund[j];
                item.EST_FORM_ID_REFUND = lstEstIdRefund[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且扣款原因為" + item.REASON + "且請款對象為" + item.CONTRACT_ID_FOR_REFUND + "且請款單號為" + item.EST_FORM_ID_REFUND);
                lstItem.Add(item);
            }
            int i = service.RefreshRefund(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新代付扣回資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }

        //取得未核准的估驗單
        public string getEstNoOfUnApproval()
        {
            logger.Debug("get EST No of un Approval By contractid:" + Request["contractid"]);
            string contractid = Request["contractid"];
            string UnApproval = null;
            UnApproval = service.getEstNoByContractId(contractid);
            return UnApproval;
        }

        //工地費用功能
        #region 工地費用
        public ActionResult Expense(string id)
        {
            logger.Info("Access to Expense Page !!");
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            List<FIN_SUBJECT> Subject = null;
            Subject = service.getSubjectOfExpense4Site();
            ViewData["items"] = JsonConvert.SerializeObject(Subject);
            return View();
        }

        public ActionResult SearchSubject()
        {
            //取得使用者勾選品項ID
            logger.Info("item_list:" + Request["subject"]);
            string[] lstItemId = Request["subject"].ToString().Split(',');
            logger.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                logger.Info("item_list return No.:" + lstItemId[i]);
            }
            ViewBag.projectid = Request["projectid"];
            TND_PROJECT p = service.getProjectById(Request["projectid"]);
            ViewBag.projectName = p.PROJECT_NAME;
            List<FIN_SUBJECT> SubjectChecked = null;
            SubjectChecked = service.getSubjectByChkItem(lstItemId);
            List<FIN_SUBJECT> Subject = null;
            Subject = service.getSubjectOfExpense4Site();
            ViewData["items"] = JsonConvert.SerializeObject(Subject);
            return View("Expense", SubjectChecked);
        }

        [HttpPost]
        public ActionResult AddExpense(FIN_EXPENSE_FORM ef)
        {
            string[] lstSubject = Request["subject"].Split(',');
            string[] lstAmount = Request["expense_amount"].Split(',');
            string[] lstRemark = Request["item_remark"].Split(',');
            //建立工地費用單號
            logger.Info("create new Plan Expense Form");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            ef.PROJECT_ID = Request["projectid"];
            ef.PAYMENT_DATE = Convert.ToDateTime(Request["paymentdate"]);
            ef.OCCURRED_YEAR = int.Parse(Request["occurreddate"].Substring(0, 4));
            ef.OCCURRED_MONTH = int.Parse(Request["occurreddate"].Substring(5, 2));
            ef.CREATE_DATE = DateTime.Now;
            ef.CREATE_ID = uInfo.USER_ID;
            ef.REMARK = Request["remark"];
            ef.STATUS = 10;
            string fid = service.newExpenseForm(ef);
            //建立工地費用單明細
            List<FIN_EXPENSE_ITEM> lstItem = new List<FIN_EXPENSE_ITEM>();
            for (int j = 0; j < lstSubject.Count(); j++)
            {
                FIN_EXPENSE_ITEM item = new FIN_EXPENSE_ITEM();
                item.FIN_SUBJECT_ID = lstSubject[j];
                item.ITEM_REMARK = lstRemark[j];
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Plan Expense Subject =" + item.FIN_SUBJECT_ID + "， and Amount = " + item.AMOUNT);
                item.EXP_FORM_ID = fid;
                logger.Debug("Item EX form id =" + item.EXP_FORM_ID);
                lstItem.Add(item);
            }
            int i = service.AddExpenseItems(lstItem);
            logger.Debug("Item Count =" + i);
            return RedirectToAction("SingleEXPForm", "CashFlow", new { id = fid });
        }

        public ActionResult SiteBudget(string id)
        {
            logger.Info("Access to Site Budget Page, And Project Id = " + id);
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得工地費用預算資料
            var priId = service.getSiteBudgetById(id);
            ViewBag.budgetdata = priId;
            List<ExpenseBudgetSummary> FirstYearBudget = null;
            List<ExpenseBudgetSummary> SecondYearBudget = null;
            List<ExpenseBudgetSummary> ThirdYearBudget = null;
            List<ExpenseBudgetSummary> FourthYearBudget = null;
            List<ExpenseBudgetSummary> FifthYearBudget = null;
            ExpenseBudgetSummary Amt = null;
            if (null != priId && priId != "")
            {
                //取得已寫入之工地費用預算資料
                FirstYearBudget = service.getFirstYearBudgetByProject(id);
                SecondYearBudget = service.getSecondYearBudgetByProject(id);
                ThirdYearBudget = service.getThirdYearBudgetByProject(id);
                FourthYearBudget = service.getFourthYearBudgetByProject(id);
                FifthYearBudget = service.getFifthYearBudgetByProject(id);
                ViewBag.FirstYear = service.getFirstYearOfSiteBudgetById(id);
                ViewBag.SecondYear = service.getSecondYearOfSiteBudgetById(id);
                ViewBag.ThirdYear = service.getThirdYearOfSiteBudgetById(id);
                ViewBag.FourthYear = service.getFourthYearOfSiteBudgetById(id);
                ViewBag.FifthYear = service.getFifthYearOfSiteBudgetById(id);
                Amt = service.getTotalSiteBudgetAmount(id);
                TempData["TotalAmt"] = Amt.TOTAL_BUDGET;
                SiteBudgetModels viewModel = new SiteBudgetModels();
                viewModel.firstYear = FirstYearBudget;
                viewModel.secondYear = SecondYearBudget;
                viewModel.thirdYear = ThirdYearBudget;
                viewModel.fourthYear = FourthYearBudget;
                viewModel.fifthYear = FifthYearBudget;
                return View(viewModel);
            }
            return View();
        }

        //更新工地費用預算
        #region 第1年度
        public String UpdateSiteBudgetOfFirstYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 1.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt1").Split(',');
            string[] lst2 = form.Get("febAmt1").Split(',');
            string[] lst3 = form.Get("marAmt1").Split(',');
            string[] lst4 = form.Get("aprAmt1").Split(',');
            string[] lst5 = form.Get("mayAmt1").Split(',');
            string[] lst6 = form.Get("junAmt1").Split(',');
            string[] lst7 = form.Get("julAmt1").Split(',');
            string[] lst8 = form.Get("augAmt1").Split(',');
            string[] lst9 = form.Get("sepAmt1").Split(',');
            string[] lst10 = form.Get("octAmt1").Split(',');
            string[] lst11 = form.Get("novAmt1").Split(',');
            string[] lst12 = form.Get("decAmt1").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["firstYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["firstYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["firstYear"]);
            return msg;
        }
        #endregion
        #region 第2年度
        public String UpdateSiteBudgetOfSecondYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 2.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt2").Split(',');
            string[] lst2 = form.Get("febAmt2").Split(',');
            string[] lst3 = form.Get("marAmt2").Split(',');
            string[] lst4 = form.Get("aprAmt2").Split(',');
            string[] lst5 = form.Get("mayAmt2").Split(',');
            string[] lst6 = form.Get("junAmt2").Split(',');
            string[] lst7 = form.Get("julAmt2").Split(',');
            string[] lst8 = form.Get("augAmt2").Split(',');
            string[] lst9 = form.Get("sepAmt2").Split(',');
            string[] lst10 = form.Get("octAmt2").Split(',');
            string[] lst11 = form.Get("novAmt2").Split(',');
            string[] lst12 = form.Get("decAmt2").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["secondYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["secondYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["secondYear"]);
            return msg;
        }
        #endregion
        #region 第3年度
        public String UpdateSiteBudgetOfThirdYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 3.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt3").Split(',');
            string[] lst2 = form.Get("febAmt3").Split(',');
            string[] lst3 = form.Get("marAmt3").Split(',');
            string[] lst4 = form.Get("aprAmt3").Split(',');
            string[] lst5 = form.Get("mayAmt3").Split(',');
            string[] lst6 = form.Get("junAmt3").Split(',');
            string[] lst7 = form.Get("julAmt3").Split(',');
            string[] lst8 = form.Get("augAmt3").Split(',');
            string[] lst9 = form.Get("sepAmt3").Split(',');
            string[] lst10 = form.Get("octAmt3").Split(',');
            string[] lst11 = form.Get("novAmt3").Split(',');
            string[] lst12 = form.Get("decAmt3").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["thirdYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["thirdYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["thirdYear"]);
            return msg;
        }
        #endregion
        #region 第4年度
        public String UpdateSiteBudgetOfFourthYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 4.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt4").Split(',');
            string[] lst2 = form.Get("febAmt4").Split(',');
            string[] lst3 = form.Get("marAmt4").Split(',');
            string[] lst4 = form.Get("aprAmt4").Split(',');
            string[] lst5 = form.Get("mayAmt4").Split(',');
            string[] lst6 = form.Get("junAmt4").Split(',');
            string[] lst7 = form.Get("julAmt4").Split(',');
            string[] lst8 = form.Get("augAmt4").Split(',');
            string[] lst9 = form.Get("sepAmt4").Split(',');
            string[] lst10 = form.Get("octAmt4").Split(',');
            string[] lst11 = form.Get("novAmt4").Split(',');
            string[] lst12 = form.Get("decAmt4").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["fourthYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["fourthYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["fourthYear"]);
            return msg;
        }
        #endregion
        #region 第5年度
        public String UpdateSiteBudgetOfFifthYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 4.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt5").Split(',');
            string[] lst2 = form.Get("febAmt5").Split(',');
            string[] lst3 = form.Get("marAmt5").Split(',');
            string[] lst4 = form.Get("aprAmt5").Split(',');
            string[] lst5 = form.Get("mayAmt5").Split(',');
            string[] lst6 = form.Get("junAmt5").Split(',');
            string[] lst7 = form.Get("julAmt5").Split(',');
            string[] lst8 = form.Get("augAmt5").Split(',');
            string[] lst9 = form.Get("sepAmt5").Split(',');
            string[] lst10 = form.Get("octAmt5").Split(',');
            string[] lst11 = form.Get("novAmt5").Split(',');
            string[] lst12 = form.Get("decAmt5").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["fourthYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["fifthYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["fifthYear"]);
            return msg;
        }
        #endregion

        //工地費用下載與上傳操作頁面
        public ActionResult SiteBudgetOperation(string id)
        {
            logger.Info("Access to Site Budget Page, And Project Id = " + id);
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            return View();
        }
        /// <summary>
        /// 下載工地費用預算填寫表
        /// </summary>
        public void downLoadSiteBudgetForm()
        {
            string projectid = Request["projectid"];
            PlanService ps = new PlanService();
            ps.getProjectId(projectid);
            if (null != ps.budgetTable)
            {
                SiteBudgetFormToExcel poi = new SiteBudgetFormToExcel();
                //檔案位置
                string fileLocation = poi.exportExcel(ps.budgetTable);
                //檔案名稱 HttpUtility.UrlEncode預設會以UTF8的編碼系統進行QP(Quoted-Printable)編碼，可以直接顯示的7 Bit字元(ASCII)就不用特別轉換。
                string filename = HttpUtility.UrlEncode(Path.GetFileName(fileLocation));
                Response.Clear();
                Response.Charset = "utf-8";
                Response.ContentType = "text/xls";
                Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", filename));
                ///"\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_ID + ".xlsx"
                Response.WriteFile(fileLocation);
                Response.End();
            }
        }

        //上傳工地費用預算
        [HttpPost]
        public ActionResult uploadSiteBudgetTable(HttpPostedFileBase fileBudget1, HttpPostedFileBase fileBudget2, HttpPostedFileBase fileBudget3,
            HttpPostedFileBase fileBudget4, HttpPostedFileBase fileBudget5)
        {
            string projectid = Request["projectid"];
            string message = "";
            logger.Info("Upload plan Budget of Site Expenses for projectid=" + projectid);
            try
            {
                //檔案變數名稱需要與前端畫面對應
                #region 預算第1年度
                //工地費用:預算第1年度
                if (null != fileBudget1 && fileBudget1.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileBudget1.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget1.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget1.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget1.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> lstSiteBudget = budgetservice.ConvertDataForSiteBudget1(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 1.ToString();
                    logger.Info("Delete PALN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add First Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(lstSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 預算第2年度
                //工地費用:預算第2年度

                if (null != fileBudget2 && fileBudget2.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileBudget1.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget1.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget1.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget1.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> _2ndSiteBudget = budgetservice.ConvertDataForSiteBudget2(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 2.ToString();
                    logger.Info("Delete PALN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add Second Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(_2ndSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 預算第3年度
                //工地費用:預算第3年度

                if (null != fileBudget3 && fileBudget3.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileBudget1.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget1.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget1.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget1.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> _3rdSiteBudget = budgetservice.ConvertDataForSiteBudget3(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 3.ToString();
                    logger.Info("Delete PALN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add Second Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(_3rdSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 預算第4年度
                //工地費用:預算第4年度

                if (null != fileBudget4 && fileBudget4.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileBudget1.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget1.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget1.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget1.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> _4thSiteBudget = budgetservice.ConvertDataForSiteBudget3(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 4.ToString();
                    logger.Info("Delete PALN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add Second Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(_4thSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 預算第5年度
                //工地費用:預算第5年度

                if (null != fileBudget5 && fileBudget5.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileBudget1.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget1.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget1.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget1.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> _5thSiteBudget = budgetservice.ConvertDataForSiteBudget5(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 5.ToString();
                    logger.Info("Delete PLAN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add Second Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(_5thSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
                message = ex.Message;
            }
            TempData["result"] = message;
            return RedirectToAction("SiteBudget/" + projectid);
        }
        #endregion

        public ActionResult SiteExpSummary(string id)
        {
            logger.Info("Access to Site Expense and Budget Summary Page，Project Id =" + id);
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.searchKey = Request["searchKey"];
            if (null != Request["searchKey"] && Request["searchKey"] != "")
            {
                int budgetYear = 0;
                int targetYear = 0;
                int targetMonth = 0;
                int sequence = 0;
                bool isCum = false;
                if (Request["searchKey"] == "S")
                {
                    budgetYear = service.getYearOfSiteExpenseById(Request["projectid"], int.Parse(Request["yearSequence"]));
                    targetYear = budgetYear;
                    int firstYear = service.getFirstYearOfPlanById(Request["projectid"]);
                    if (budgetYear == 0)
                    {
                        sequence = int.Parse(Request["yearSequence"]);
                        targetYear = firstYear + sequence - 1;
                    }
                }
                else
                {
                    string[] date = Request["date"].Split('/');
                    targetYear = int.Parse(date[0]);
                    targetMonth = int.Parse(date[1]);
                    isCum = true;
                }
                List<ExpenseBudgetSummary> ExpBudget = null;
                List<ExpenseBudgetByMonth> BudgetByMonth = null;
                List<ExpensetFromOPByMonth> ExpenseByMonth = null;
                ExpenseBudgetSummary Amt = null;
                ExpenseBudgetSummary ExpAmt = null;
                ExpenseBudgetModel viewModel = new ExpenseBudgetModel();
                ExpBudget = service.getSiteExpBudgetSummaryBySeqYear(Request["projectid"], sequence, targetYear, targetMonth, isCum);
                BudgetByMonth = service.getSiteExpBudgetOfMonth(Request["projectid"], sequence, targetYear, targetMonth, isCum);
                ExpenseByMonth = service.getSiteExpensetOfMonth(Request["projectid"], targetYear, targetMonth, isCum);
                Amt = service.getSiteBudgetAmountById(Request["projectid"]);
                ExpAmt = service.getTotalSiteExpAmountById(Request["projectid"], targetYear, targetMonth, isCum);
                viewModel.summary = ExpBudget;
                viewModel.budget = BudgetByMonth;
                viewModel.expense = ExpenseByMonth;
                TempData["TotalAmt"] = Amt.TOTAL_BUDGET;
                TempData["TotalExpAmt"] = ExpAmt.CUM_YEAR_AMOUNT;
                TempData["targetYear"] = targetYear;
                TempData["targetMonth"] = targetMonth;
                TempData["yearSequence"] = sequence;
                return View(viewModel);
            }
            return View();
        }

        public ActionResult SearchSiteExpSummary()
        {
            logger.Info("Access to Site Expense and Budget Summary Page，Project Id =" + Request["projectid"]);
            ViewBag.projectid = Request["projectid"];
            ViewBag.projectName = Request["projectName"];
            ViewBag.searchKey = Request["searchKey"];
            logger.Debug("search date = " + Request["date"]);
            int budgetYear = 0;
            int targetYear = 0;
            int targetMonth = 0;
            int sequence = 0;
            bool isCum = false;
            if (Request["searchKey"] == "S")
            {
                budgetYear = service.getYearOfSiteExpenseById(Request["projectid"], int.Parse(Request["yearSequence"]));
                targetYear = budgetYear;
                int firstYear = service.getFirstYearOfPlanById(Request["projectid"]);
                if (budgetYear == 0)
                {
                    sequence = int.Parse(Request["yearSequence"]);
                    targetYear = firstYear + sequence - 1;
                }
            }
            else
            {
                string[] date = Request["date"].Split('/');
                targetYear = int.Parse(date[0]);
                targetMonth = int.Parse(date[1]);
                isCum = true;
            }
            List<ExpenseBudgetSummary> ExpBudget = null;
            List<ExpenseBudgetByMonth> BudgetByMonth = null;
            List<ExpensetFromOPByMonth> ExpenseByMonth = null;
            ExpenseBudgetSummary Amt = null;
            ExpenseBudgetSummary ExpAmt = null;
            ExpenseBudgetModel viewModel = new ExpenseBudgetModel();
            if (null != Request["searchKey"] && Request["searchKey"] != "")
            {
                ExpBudget = service.getSiteExpBudgetSummaryBySeqYear(Request["projectid"], sequence, targetYear, targetMonth, isCum);
                BudgetByMonth = service.getSiteExpBudgetOfMonth(Request["projectid"], sequence, targetYear, targetMonth, isCum);
                ExpenseByMonth = service.getSiteExpensetOfMonth(Request["projectid"], targetYear, targetMonth, isCum);
                Amt = service.getSiteBudgetAmountById(Request["projectid"]);
                ExpAmt = service.getTotalSiteExpAmountById(Request["projectid"], targetYear, targetMonth, isCum);
                viewModel.summary = ExpBudget;
                viewModel.budget = BudgetByMonth;
                viewModel.expense = ExpenseByMonth;
                TempData["TotalAmt"] = Amt.TOTAL_BUDGET;
                TempData["TotalExpAmt"] = ExpAmt.CUM_YEAR_AMOUNT;
                TempData["targetYear"] = targetYear;
                TempData["targetMonth"] = targetMonth;
                TempData["yearSequence"] = sequence;
                return View("SiteExpSummary", viewModel);
            }
            return View("SiteExpSummary");
        }

        /// <summary>
        ///  工地費用預算執行彙整表
        /// </summary>
        public void downLoadSiteExpenseSummary()
        {
            string projectid = Request["projectid"];
            int targetYear = int.Parse(Request["targetYear"]);
            int sequence = 0;
            int targetMonth = 0;
            bool isCum = false;
            if (null != Request["isCum"] && Request["isCum"] == "Y")
            {
                isCum = true;
                targetMonth = int.Parse(Request["targetMonth"]);
            }
            else
            {
                sequence = int.Parse(Request["sequence"]);
            }
            SiteExpSummaryToExcel poi = new SiteExpSummaryToExcel();
            //檔案位置
            string fileLocation = poi.exportExcel(projectid, sequence, targetYear, targetMonth, isCum);
            //檔案名稱 HttpUtility.UrlEncode預設會以UTF8的編碼系統進行QP(Quoted-Printable)編碼，可以直接顯示的7 Bit字元(ASCII)就不用特別轉換。
            string filename = HttpUtility.UrlEncode(Path.GetFileName(fileLocation));
            Response.Clear();
            Response.Charset = "utf-8";
            Response.ContentType = "text/xls";
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", filename));
            ///"\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_ID + ".xlsx"
            Response.WriteFile(fileLocation);
            Response.End();
        }

        //上傳業主計價項目
        public string uploadVAItem4Owner(HttpPostedFileBase fileValuation)
        {
            string projectId = Request["id"];
            logger.Info("Upload valuation item for owner, project id =" + projectId);
            if (null != fileValuation && fileValuation.ContentLength != 0)
            {
                try
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileValuation.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileValuation.FileName);
                    var path = Path.Combine(ContextService.strUploadPath, fileName);
                    logger.Info("save excel file:" + path);
                    fileValuation.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileValuation.FileName);
                    VAItem4OwnerToExcel ownerservice = new VAItem4OwnerToExcel();
                    ownerservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_VALUATION_4OWNER> lstVAItem = ownerservice.ConvertDataForVAOfOwner(projectId);
                    //2.3
                    logger.Info("Delete PLAN_VALUATION_4OWNER By Project Id");
                    service.delVAItemOfOwnerById(projectId);
                    //2.4 
                    logger.Info("Add All PLAN_VALUATION_4OWNER to DB");
                    service.refreshVAItemOfOwner(lstVAItem);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
                    return ex.Message;
                }
            }
            if (service.strMessage != null)
            {
                return service.strMessage;
            }
            else
            {
                return "匯入成功!!";
            }
        }
        //業主估驗計價
        public ActionResult Valuation4Owner(string id)
        {
            logger.Info("Access to Valuation For Owner Page，Project Id =" + id);
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            var priId = service.getVAOfOwnerById(id);
            if (null != priId && priId != "")
            {
                List<RevenueFromOwner> lstVAItem = null;
                lstVAItem = service.getVAItemOfOwnerById(id);
                ViewBag.SearchResult = "共取得" + lstVAItem.Count + "筆資料";
                //轉成Json字串
                ViewData["items"] = JsonConvert.SerializeObject(lstVAItem);
            }
            return View();
        }

        public string getVAItem(string itemid)
        {
            logger.Info("get valuatio item by item no =" + itemid);
            string[] key = itemid.Split('+');
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getVAItem(key[0], key[1]));
            logger.Info("valuatio item  info=" + itemJson);
            return itemJson;
        }
        
        public String updateVAItem(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "更新成功!!";

            PLAN_VALUATION_4OWNER item = new PLAN_VALUATION_4OWNER();
            item.PROJECT_ID = form["project_id"];
            item.ITEN_NO = form["item_no"];
            item.ITEM_DESC = form["item_desc"];
            item.REMARK = form["item_remark"];
            try
            {
                item.ITEM_VALUATION_RATIO = decimal.Parse(form["item_valuation_ratio"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.ITEN_NO + " not valuation_ratio:" + ex.Message);
            }
            SYS_USER loginUser = (SYS_USER)Session["user"];
            item.MODIFY_ID = loginUser.USER_ID;
            item.MODIFY_DATE = DateTime.Now;
            int i = 0;
            i = service.updateVAItem(item);
            if (i == 0) { msg = service.message; }
            return msg;
        }
        //新增業主計價單
        public ActionResult AddVAForm(PLAN_VALUATION_FORM vf)
        {
            //取得專案編號
            logger.Info("Project Id:" + Request["id"]);
            //取得專案名稱
            logger.Info("Project Name:" + Request["projectName"]);
            //取得使用者勾選品項ID
            logger.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            logger.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                logger.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] Amt = Request["evaluated_amount"].Split(',');
            List<string> lstAmt = new List<string>();
            var m = 0;
            for (m = 0; m < Amt.Count(); m++)
            {
                if (Amt[m] != "" && null != Amt[m])
                {
                    lstAmt.Add(Amt[m]);
                }
            }
            //建立計價單
            logger.Info("create new Valuation Form");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            vf.PROJECT_ID = Request["id"];
            vf.CREATE_ID = u.USER_ID;
            vf.CREATE_DATE = DateTime.Now;
            PLAN_VALUATION_FORM_ITEM item = new PLAN_VALUATION_FORM_ITEM();
            string vaid = service.newVA(Request["id"], vf, lstItemId);
            List<PLAN_VALUATION_FORM_ITEM> lstItem = new List<PLAN_VALUATION_FORM_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_VALUATION_FORM_ITEM items = new PLAN_VALUATION_FORM_ITEM();
                items.ITEM_NO = lstItemId[j];
                items.ITEM_VALUATION_AMOUNT = decimal.Parse(lstAmt[j]);
                logger.Debug("Item No=" + items.ITEM_NO + ", Amt =" + items.ITEM_VALUATION_AMOUNT);
                lstItem.Add(items);
            }
            int k = service.refreshVA(vaid, vf, lstItem);
            return Redirect("SingleVA?id=" + vaid);
        }
    }
}