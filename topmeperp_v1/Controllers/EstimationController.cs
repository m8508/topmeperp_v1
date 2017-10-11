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

        public ActionResult AddEST(string id)
        {
            logger.Info("create EST form id process!  Contract id=" + id);
            //新增估驗單編號
            string formid = service.getEstNo();
            return RedirectToAction("ContractItems", "Estimation", new { id = id, formid = formid });
        }

        public ActionResult ContractItems(string id, string formid)
        {
            logger.Info("Access To Contract Item By Contract Id =" + id);
            ViewBag.projectId = id.Substring(1, 6).Trim();
            TND_PROJECT p = service.getProjectById(id.Substring(1, 6).Trim());
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.wage = id.Substring(0, 1).Trim();
            ContractModels contract = new ContractModels();
            ViewBag.formid = formid;
            ViewBag.keyid = id; //使用供應商名稱的contractid
            //取得合約金額與供應商名稱,採購項目等資料
            if (ViewBag.wage == "")
            {
                plansummary lstContract = service.getPlanContract4Est(id.Substring(1).Trim());
                ViewBag.supplier = lstContract.SUPPLIER_ID;
                ViewBag.formname = lstContract.FORM_NAME;
                ViewBag.amount = lstContract.MATERIAL_COST;
                ViewBag.contractid = lstContract.CONTRACT_ID; //使用供應商編號的contractid
                PLAN_PAYMENT_TERMS payment = service.getPaymentTerm(lstContract.CONTRACT_ID);
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
                    ViewBag.type = 'M'; // 材料合約(含設備)
                }
                ViewBag.estCount = service.getEstCountById(lstContract.CONTRACT_ID);
                ViewBag.paymentTerms = service.getTermsByContractId(lstContract.CONTRACT_ID);
                var balance = service.getBalanceOfRefundById(lstContract.CONTRACT_ID);
                if (balance > 0)
                {
                    TempData["balance"] = "本合約目前尚有 " + string.Format("{0:C0}", balance) + "的代付支出款項，仍未扣回!";
                }
            }
            else
            {
                plansummary lstWageContract = service.getPlanContractOfWage4Est(id.Substring(1).Trim());
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
            ViewBag.paymentkey = ViewBag.formid + ViewBag.contractid;
            List<EstimationForm> lstContractItem = null;
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
        public ActionResult SaveEst(FormCollection f)
        {
            //取得估驗單編號
            logger.Info("EST form Id:" + Request["formid"]);
            logger.Info("Project Id:" + Request["contractid"].Substring(0, 6).Trim());
            logger.Info("get EST No of un Approval By contractid:" + Request["contractid"]);
            string contractid = Request["contractid"];
            string UnApproval = null;
            UnApproval = service.getEstNoByContractId(contractid);
            if (UnApproval != null && "" != UnApproval)
            {
                TempData["result"] = "目前尚有未核准的估驗單，估驗單編號為" + UnApproval + "，待此單核准後再新增估驗單!";
                return RedirectToAction("Valuation", "Estimation", new { id = Request["contractid"].Substring(0, 6).Trim() });
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
            logger.Info("Project Id:" + Request["contractid"].Substring(0, 6).Trim());
            string contractid = Request["contractid"];
            string UnApproval = null;
            UnApproval = service.getEstNoByContractId(contractid);
            if (UnApproval != null && "" != UnApproval)
            {
                TempData["result"] = "目前尚有未核准的估驗單，估驗單編號為" + UnApproval + "，待此單核准後再新增估驗單!";
                return RedirectToAction("Valuation", "Estimation", new { id = Request["contractid"].Substring(0, 6).Trim() });
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
            logger.Info("ContractId:" + Request["keyid"]);
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
                int m = service.UpdateRetentionAmountById(Request["formid"] + Request["contractid"]);
                System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                string itemJson = objSerializer.Serialize(service.getDetailsPayById(Request["formid"] + Request["contractid"]));
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
            PaymentDetailsFunction lstSummary = service.getDetailsPayById(id + singleForm.planEST.CONTRACT_ID);
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
        public ActionResult OtherPayment(string id)
        {
            logger.Info("Access To Other Payment By EST Form Id =" + id);
            ViewBag.keyId = id;
            ViewBag.projectId = id.Substring(9, 6).Trim();
            TND_PROJECT p = service.getProjectById(id.Substring(9, 6).Trim());
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = id.Substring(9).Trim();
            ViewBag.formname = id.Substring(22).Trim();
            ViewBag.formid = id.Substring(0, 9).Trim();
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
            int m = service.UpdateRetentionAmountById(form["estid"] + form["contractid"]);
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
        public ActionResult AdvancePayment(string id)
        {
            logger.Info("Access To Advance Payment By EST Form Id =" + id);
            ViewBag.keyId = id;
            ViewBag.projectId = id.Substring(9, 6).Trim();
            TND_PROJECT p = service.getProjectById(id.Substring(9, 6).Trim());
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = id.Substring(9).Trim();
            ViewBag.formid = id.Substring(0, 9).Trim();
            PLAN_PAYMENT_TERMS payment = service.getPaymentTerm(id.Substring(9).Trim());
            ViewBag.advancePaymentRatio = payment.PAYMENT_ADVANCE_RATIO;
            AdvancePaymentFunction advancePay = service.getAdvancePayById(id);
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
        public ActionResult Invoice(string id)
        {
            logger.Info("Access To Invoice By EST Form Id =" + id);
            ViewBag.keyId = id;
            ViewBag.projectId = id.Substring(9, 6).Trim();
            TND_PROJECT p = service.getProjectById(id.Substring(9, 6).Trim());
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.formname = id.Substring(22).Trim();
            ViewBag.contractid = id.Substring(9).Trim();
            ViewBag.formid = id.Substring(0, 9).Trim();
            TND_SUPPLIER lstSupplier = service.getSupplierInfo(id.Substring(15, 7).Trim());
            ViewBag.supplier = lstSupplier.COMPANY_NAME;
            ViewBag.companyid = lstSupplier.COMPANY_ID;
            ContractModels singleForm = new ContractModels();
            service.getESTByEstId(id.Substring(0, 9).Trim());
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
        public ActionResult RePayment(string id)
        {
            logger.Info("Access To RePayment By EST Form Id =" + id);
            ViewBag.keyId = id;
            ViewBag.projectId = id.Substring(9, 6).Trim();
            TND_PROJECT p = service.getProjectById(id.Substring(9, 6).Trim());
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = id.Substring(9).Trim();
            ViewBag.formname = id.Substring(22).Trim();
            ViewBag.formid = id.Substring(0, 9).Trim();
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
        public ActionResult ChooseSupplier(string id)
        {
            logger.Info("Access To RePayment By EST Form Id =" + id);
            ViewBag.keyId = id;
            ViewBag.contractid = id.Substring(9).Trim();
            ViewBag.formid = id.Substring(0, 9).Trim();
            List<RePaymentFunction> lstSupplier = null;
            lstSupplier = service.getSupplierOfContractByPrjId(id.Substring(9, 6).Trim());
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstSupplier.Count;
            logger.Debug("supplier record =" + ViewBag.key + "筆");
            return View(lstSupplier);
        }

        public ActionResult AddRePayment(string id)
        {
            //取得估驗單編號
            logger.Info("EST Form Id:" + Request["formid"]);
            //取得合約編號
            logger.Info("Contract Id:" + Request["contractid"]);
            //取得KEY_ID
            logger.Info("Key Id:" + Request["keyid"]);
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
            return Redirect("RePayment?id=" + Request["keyid"]);
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
        public ActionResult Refund(string id)
        {
            logger.Info("Access To Refund By EST Form Id =" + id);
            ViewBag.keyId = id;
            ViewBag.projectId = id.Substring(9, 6).Trim();
            TND_PROJECT p = service.getProjectById(id.Substring(9, 6).Trim());
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = id.Substring(9).Trim();
            ViewBag.formname = id.Substring(22).Trim();
            ViewBag.formid = id.Substring(0, 9).Trim();
            List<RePaymentFunction> lstOtherPayItem = null;
            List<RePaymentFunction> lstRefundItem = null;
            lstOtherPayItem = service.getRefundById(id);
            lstRefundItem = service.getRefundOfSupplierById(id);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstOtherPayItem.Count;
            logger.Debug("this repayment record =" + ViewBag.key + "筆");
            ViewData["items"] = JsonConvert.SerializeObject(lstOtherPayItem);
            logger.Debug(ViewData["items"]);
            return View(lstRefundItem);
        }

        //代付支出-選商
        public ActionResult ChooseSupplierOfRefund(string id)
        {
            logger.Info("Access To Refund By EST Form Id =" + id);
            ViewBag.keyId = id;
            ViewBag.contractid = id.Substring(9).Trim();
            ViewBag.formid = id.Substring(0, 9).Trim();
            List<RePaymentFunction> lstSupplier = null;
            lstSupplier = service.getSupplierOfContractRefundById(id.Substring(9).Trim());
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstSupplier.Count;
            logger.Debug("supplier record =" + ViewBag.key + "筆");
            return View(lstSupplier);
        }

        public ActionResult AddRefund(string id)
        {
            //取得估驗單編號
            logger.Info("EST Form Id:" + Request["formid"]);
            //取得合約編號
            logger.Info("Contract Id:" + Request["contractid"]);
            //取得KEY_ID
            logger.Info("Key Id:" + Request["keyid"]);
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
            return Redirect("Refund?id=" + Request["keyid"]);
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

        
    }
}