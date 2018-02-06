﻿using log4net;
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
            List<ProjectList> lstProject = SearchProjectByName("", "專案執行");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";

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

        //上傳得標後標單內容(用於標單內容有異動時)_2017/9/8
        [HttpPost]
        public ActionResult uploadPlanItem(TND_PROJECT prj, HttpPostedFileBase file)
        {
            //1.取得專案編號
            string projectid = Request["projectid"];
            logger.Info("Upload plan items for projectid=" + projectid);
            string message = "";
            TND_PROJECT p = service.getProject(projectid);
            if (null != file && file.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + file.FileName);
                //2.1 將上傳檔案存檔
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                file.SaveAs(path);
                //2.2 解析Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                TnderProject tndService = new TnderProject();
                tndService.project = p;
                poiservice.InitializeWorkbook(path);
                poiservice.ConvertDataForPlan(projectid);
                //2.3 記錄錯誤訊息
                message = message + "得標標單品項:共" + poiservice.lstPlanItem.Count + "筆資料，";
                message = message + "<a target=\"_blank\" href=\"/Plan/ManagePlanItem?id=" + projectid + "\"> 標單明細檢視畫面單</a><br/>" + poiservice.errorMessage;
                //        < button type = "button" class="btn btn-primary" onclick="location.href='@Url.Action("ManagePlanItem","Plan", new { id = @Model.tndProject.PROJECT_ID})'; ">標單明細</button>
                //2.4
                logger.Info("Delete PLAN_ITEM By Project ID");
                tndService.delAllItemByPlan();
                //2.5
                logger.Info("Add All PLAN_ITEM to DB");
                tndService.refreshPlanItem(poiservice.lstPlanItem);
            }
            TempData["result"] = message;
            //PlanService ps = new PlanService();
            //var priId = ps.getBudgetById(projectid);
            //ViewBag.budgetdata = priId;
            //if (null != priId) { int k = service.updateBudgetToPlanItem(projectid); }
            return Redirect("ManagePlanItem?id=" + projectid);
        }
        /// <summary>
        /// 設定標單品項查詢條件
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public ActionResult ManagePlanItem(string id)
        {
            //傳入專案編號，
            PurchaseFormService service = new PurchaseFormService();
            logger.Info("start project id=" + id);

            //取得專案基本資料fc
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
        /// <summary>
        /// 取得標單明細資料
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public ActionResult ShowPlanItems(FormCollection form)
        {
            PurchaseFormService service = new PurchaseFormService();
            logger.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"]);
            logger.Debug("Exception check=" + Request["chkEx"]);
            List<PlanItem4Map> lstItems = service.getPlanItem(Request["chkEx"], Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"], Request["supplier"], Request["selDelFlag"]);
            ViewBag.Result = "共" + lstItems.Count + "筆資料";
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
            return PartialView(lstItems);
        }
        public string getPlanItem(string itemid)
        {
            PurchaseFormService service = new PurchaseFormService();
            logger.Info("get plan item by id=" + itemid);
            //PLAN_ITEM  item = service.getPlanItem.getUser(userid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getPlanItem(itemid));
            logger.Info("plan item  info=" + itemJson);
            return itemJson;
        }
        public String addPlanItem(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "更新成功!!";

            PLAN_ITEM item = new PLAN_ITEM();
            item.PROJECT_ID = form["project_id"];
            item.PLAN_ITEM_ID = form["plan_item_id"];
            item.ITEM_ID = form["item_id"];
            item.ITEM_DESC = form["item_desc"];
            item.ITEM_UNIT = form["item_unit"];
            try
            {
                item.ITEM_QUANTITY = decimal.Parse(form["item_quantity"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not quattity:" + ex.Message);
            }
            try
            {
                item.ITEM_FORM_QUANTITY = decimal.Parse(form["item_form_quantity"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not form quattity:" + ex.Message);
            }
            try
            {
                item.ITEM_UNIT_PRICE = decimal.Parse(form["item_unit_price"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not unit price:" + ex.Message);
            }
            try
            {
                item.ITEM_UNIT_COST = decimal.Parse(form["item_unit_cost"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not unit cost:" + ex.Message);
            }
            item.ITEM_REMARK = form["item_remark"];
            if (form["type_code_1"].Trim() != "")
            {
                item.TYPE_CODE_1 = form["type_code_1"];
            }
            else
            {
                item.TYPE_CODE_1 = null;
            }

            if (form["type_code_2"].Trim() != "")
            {
                item.TYPE_CODE_2 = form["type_code_2"];
            }
            else
            {
                item.TYPE_CODE_2 = null;
            }

            item.SYSTEM_MAIN = form["system_main"];
            item.SYSTEM_SUB = form["system_sub"];
            item.DEL_FLAG = form["selDelFlag"];
            try
            {
                item.EXCEL_ROW_ID = long.Parse(form["excel_row_id"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not exce row id:" + ex.Message);
            }

            SYS_USER loginUser = (SYS_USER)Session["user"];
            item.MODIFY_USER_ID = loginUser.USER_ID;
            item.MODIFY_DATE = DateTime.Now;
            PurchaseFormService service = new PurchaseFormService();
            int i = 0;
            string strFlag = form["flag"].Trim();
            if (strFlag.Equals("addAfter"))
            {
                i = service.addPlanItemAfter(item);
            }
            else
            {
                i = service.updatePlanItem(item);
            }

            if (i == 0) { msg = service.message; }
            return msg;
        }
        /// <summary>
        /// Project_item 註記刪除
        /// </summary>
        /// <param name="itemid"></param>
        /// <returns></returns>
        public String delPlanItem(string itemid)
        {
            PurchaseFormService service = new PurchaseFormService();
            string msg = "更新成功!!";
            logger.Info("del plan item by id=" + itemid);
            int i = service.changePlanItem(itemid, "Y");
            return msg + "(" + i + ")";
        }
        /// <summary>
        /// 上載現有得標標單資料
        /// </summary>
        public string uploadPlanItem(HttpPostedFileBase file)
        {
            logger.Debug("ProjectID=" + Request["id"] + ",Upload PlanItem=" + file.FileName);
            return "TEST";
        }

        //取得業主合約金額
        public ActionResult ContractForOwner(string id)
        {
            //傳入專案編號，
            logger.Info("start project id=" + id);
            //取得專案基本資料
            ViewBag.id = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得業主合約金額
            PlanRevenue contractAmount = service.getPlanRevenueById(id);
            ViewBag.Amount = (null == contractAmount.PLAN_REVENUE ? 0 : contractAmount.PLAN_REVENUE);
            ViewBag.contractid = "Owner" + contractAmount.CONTRACT_ID;
            int i = service.addContractId4Owner(id);
            return View();
        }

        public ActionResult Budget(string id)
        {
            logger.Info("budget info for projectid=" + id);
            ViewBag.projectid = id;
            TnderProject service = new TnderProject();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.wageAmount = p.WAGE_MULTIPLIER;
            if (p.WAGE_MULTIPLIER == null || p.WAGE_MULTIPLIER == 0)
            {
                TempData["wagePrice"] = "單日工資金額尚未輸入，工資成本無法計算!!";
            }
            //取得直接成本資料
            PlanService ps = new PlanService();
            var priId = ps.getBudgetById(id);
            ViewBag.budgetdata = priId;
            //for exception 
            ViewBag.budget = 0;
            ViewBag.cost = 0;
            if (null == priId)
            //取得九宮格組合之直接成本資料
            {
                CostAnalysisDataService s = new CostAnalysisDataService();
                List<DirectCost> budget1 = s.getDirectCost4Budget(id);
                ViewBag.result = "共有" + (budget1.Count) + "筆資料";
                return View(budget1);
            }
            //取得已寫入之九宮格組合預算資料
            BudgetDataService bs = new BudgetDataService();
            List<DirectCost> budget2 = bs.getBudget(id);
            DirectCost totalinfo = bs.getTotalCost(id);
            ViewBag.budget = (null == totalinfo.MATERIAL_BUDGET ? 0 : totalinfo.MATERIAL_BUDGET);
            ViewBag.wagebudget = (null == totalinfo.WAGE_BUDGET ? 0 : totalinfo.WAGE_BUDGET);
            ViewBag.totalbudget = (null == totalinfo.TOTAL_BUDGET ? 0 : totalinfo.TOTAL_BUDGET);
            ViewBag.cost = (null == totalinfo.TOTAL_COST ? 0 : totalinfo.TOTAL_COST);
            ViewBag.p_cost = (null == totalinfo.TOTAL_P_COST ? 0 : totalinfo.TOTAL_P_COST);
            ViewBag.result = "共有" + budget2.Count + "筆資料";
            return View(budget2);
        }
        /// <summary>
        /// 下載預算填寫表
        /// </summary>
        public void downLoadBudgetForm()
        {
            string projectid = Request["projectid"];
            service.getProjectId(projectid);
            if (null != service.budgetTable)
            {
                BudgetFormToExcel poi = new BudgetFormToExcel();
                //檔案位置
                string fileLocation = poi.exportExcel(service.budgetTable);
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
        //上傳預算
        [HttpPost]
        public ActionResult uploadBudgetTable(HttpPostedFileBase fileBudget)
        {
            string projectid = Request["projectid"];
            logger.Info("Upload Budget Table for projectid=" + projectid);
            string message = "";
            //檢查工率乘數是否存在
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(Request["projectid"]);
            //ViewBag.projectWage = p.WAGE_MULTIPLIER;
            //if (null == ViewBag.projectWage) { throw new Exception(Request["projectid"] + "'s Wage Multiplier is not exist !!");}
            //檔案變數名稱(fileBudget)需要與前端畫面對應(view 的 file name and file id)
            if (null != fileBudget && fileBudget.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + fileBudget.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileBudget.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                fileBudget.SaveAs(path);
                //2.2 開啟Excel 檔案
                logger.Info("Parser Excel File Begin:" + fileBudget.FileName);
                BudgetFormToExcel budgetservice = new BudgetFormToExcel();
                budgetservice.InitializeWorkbook(path);
                //解析預算數量
                List<PLAN_BUDGET> lstBudget = budgetservice.ConvertDataForBudget(projectid);
                //2.3 記錄錯誤訊息
                message = budgetservice.errorMessage;
                //2.4
                logger.Info("Delete PLAN_BUDGET By Project ID");
                service.delBudgetByProject(projectid);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                logger.Info("Add All PLAN_BUDGET to DB");
                service.refreshBudget(lstBudget);
                message = message + "<br/>資料匯入完成 !!";
            }
            TempData["result"] = message;
            // 將預算寫入得標標單
            int k = service.updateBudgetToPlanItem(projectid);
            return RedirectToAction("Budget/" + projectid);
        }
        //寫入預算
        public String UpdateBudget(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            SYS_USER u = (SYS_USER)Session["user"];
            string msg = "";
            string[] lsttypecode = form.Get("code1").Split(',');
            string[] lsttypesub = form.Get("code2").Split(',');
            //string[] lstsystemmain = form.Get("systemmain").Split(',');
            //string[] lstsystemsub = form.Get("systemsub").Split(',');
            //string[] lstCost = form.Get("inputtndratio").Split(',');
            string[] lstPrice = form.Get("inputbudget").Split(',');
            string[] lstWagePrice = form.Get("inputbudget4wage").Split(',');
            List<PLAN_BUDGET> lstItem = new List<PLAN_BUDGET>();
            for (int j = 0; j < lstPrice.Count(); j++)
            {
                PLAN_BUDGET item = new PLAN_BUDGET();
                item.PROJECT_ID = form["id"];
                if (lstPrice[j].ToString() == "")
                {
                    item.BUDGET_RATIO = null;
                }
                else
                {
                    item.BUDGET_RATIO = decimal.Parse(lstPrice[j]);
                }
                if (lstWagePrice[j].ToString() == "")
                {
                    item.BUDGET_WAGE_RATIO = null;
                }
                else
                {
                    item.BUDGET_WAGE_RATIO = decimal.Parse(lstWagePrice[j]);
                }
                //if (lstCost[j].ToString() == "")
                //{
                //    item.TND_RATIO = null;
                //}
                //else
                //{
                //    item.TND_RATIO = decimal.Parse(lstCost[j]);
                //}
                logger.Info("Budget ratio =" + item.BUDGET_RATIO);
                item.TYPE_CODE_1 = lsttypecode[j];
                item.TYPE_CODE_2 = lsttypesub[j];
                //item.SYSTEM_MAIN = lstsystemmain[j];
                //item.SYSTEM_SUB = lstsystemsub[j];
                item.CREATE_ID = u.USER_ID;
                logger.Debug("Item Project id =" + item.PROJECT_ID + "且九宮格組合為" + item.TYPE_CODE_1 + item.TYPE_CODE_2 + item.SYSTEM_MAIN + item.SYSTEM_SUB);
                lstItem.Add(item);
            }
            int i = service.addBudget(lstItem);
            // 將預算寫入得標標單
            int k = service.updateBudgetToPlanItem(form["id"]);
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
        //修改預算
        public String RefreshBudget(string id, FormCollection form)
        {
            logger.Info("form:" + form.Count);
            id = Request["id"];
            SYS_USER u = (SYS_USER)Session["user"];
            string msg = "";
            string[] lsttypecode = form.Get("code1").Split(',');
            string[] lsttypesub = form.Get("code2").Split(',');
            string[] lstPrice = form.Get("inputbudget").Split(',');
            string[] lstWagePrice = form.Get("inputbudget4wage").Split(',');
            List<PLAN_BUDGET> lstItem = new List<PLAN_BUDGET>();
            for (int j = 0; j < lstPrice.Count(); j++)
            {
                PLAN_BUDGET item = new PLAN_BUDGET();
                item.PROJECT_ID = form["id"];
                if (lstPrice[j].ToString() == "")
                {
                    item.BUDGET_RATIO = null;
                }
                else
                {
                    item.BUDGET_RATIO = decimal.Parse(lstPrice[j]);
                }
                if (lstWagePrice[j].ToString() == "")
                {
                    item.BUDGET_WAGE_RATIO = null;
                }
                else
                {
                    item.BUDGET_WAGE_RATIO = decimal.Parse(lstWagePrice[j]);
                }
                logger.Info("Budget ratio =" + item.BUDGET_RATIO);
                if (lsttypecode[j].ToString() == "")
                {
                    item.TYPE_CODE_1 = "";
                }
                else
                {
                    item.TYPE_CODE_1 = lsttypecode[j];
                }
                if (lsttypesub[j].ToString() == "")
                {
                    item.TYPE_CODE_2 = "";
                }
                else
                {
                    item.TYPE_CODE_2 = lsttypesub[j];
                }
                item.MODIFY_ID = u.USER_ID;
                logger.Debug("Item Project id =" + item.PROJECT_ID + "且九宮格組合為" + item.TYPE_CODE_1 + item.TYPE_CODE_2 + item.SYSTEM_MAIN + item.SYSTEM_SUB);
                lstItem.Add(item);
            }
            int i = service.updateBudget(id, lstItem);
            int k = service.updateBudgetToPlanItem(form["id"]);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新預算資料成功，PROJECT_ID =" + form["id"];
            }

            logger.Info("Request:PROJECT_ID =" + form["id"]);
            return msg;
        }
        //複合詢價單預算
        public ActionResult BudgetForForm(string id)
        {
            logger.Info("budget for form : projectid=" + id);
            ViewBag.projectId = id;
            return View();
        }
        [HttpPost]
        public ActionResult BudgetForForm(FormCollection f)
        {

            logger.Info("projectid=" + Request["projectid"] + ",textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"] + ",formName=" + Request["formName"]);
            PurchaseFormService service = new PurchaseFormService();
            List<topmeperp.Models.PlanItem4Map> lstProject = service.getPlanItem(Request["chkEx"], Request["projectid"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formName"], Request["supplier"], "N");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            ViewBag.projectId = Request["projectid"];
            ViewBag.textCode1 = Request["textCode1"];
            ViewBag.textCode2 = Request["textCode2"];
            ViewBag.textSystemMain = Request["textSystemMain"];
            ViewBag.textSystemSub = Request["textSystemSub"];
            BudgetDataService bs = new BudgetDataService();
            DirectCost totalinfo = bs.getTotalCost(Request["projectid"]);
            DirectCost iteminfo = bs.getItemBudget(Request["projectid"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formName"]);


            ViewBag.budget = (null == totalinfo.MATERIAL_BUDGET ? 0 : totalinfo.MATERIAL_BUDGET);
            ViewBag.wagebudget = (null == totalinfo.WAGE_BUDGET ? 0 : totalinfo.WAGE_BUDGET);
            ViewBag.totalbudget = (null == totalinfo.TOTAL_BUDGET ? 0 : totalinfo.TOTAL_BUDGET);
            //ViewBag.cost =  (null == totalinfo.TOTAL_COST ? 0 : totalinfo.TOTAL_COST);
            ViewBag.itembudget = (null == iteminfo.ITEM_BUDGET ? 0 : iteminfo.ITEM_BUDGET);
            ViewBag.itemwagebudget = (null == iteminfo.ITEM_BUDGET_WAGE ? 0 : iteminfo.ITEM_BUDGET_WAGE);
            //ViewBag.itemcost = (null == iteminfo.ITEM_COST ? 0 : iteminfo.ITEM_COST); 
            return View("BudgetForForm", lstProject);
        }
        //修改Plan_Item 個別預算
        public String UpdatePlanBudget(string id, FormCollection form)
        {
            logger.Info("form:" + form.Count);
            id = Request["projectId"];
            SYS_USER u = (SYS_USER)Session["user"];
            string msg = "";
            string[] lstplanitemid = form.Get("planitemid").Split(',');
            string[] lstBudget = form.Get("budgetratio").Split(',');
            List<PLAN_ITEM> lstItem = new List<PLAN_ITEM>();
            for (int j = 0; j < lstBudget.Count(); j++)
            {
                PLAN_ITEM item = new PLAN_ITEM();
                item.PROJECT_ID = form["id"];
                if (lstBudget[j].ToString() == "")
                {
                    item.BUDGET_RATIO = null;
                }
                else
                {
                    item.BUDGET_RATIO = decimal.Parse(lstBudget[j]);
                }
                logger.Info("Budget ratio =" + item.BUDGET_RATIO);
                item.PLAN_ITEM_ID = lstplanitemid[j];
                item.MODIFY_USER_ID = u.USER_ID;
                logger.Debug("Item Project id =" + item.PROJECT_ID + "且plan item id 為" + item.PLAN_ITEM_ID + "其項目預算為" + item.BUDGET_RATIO);
                lstItem.Add(item);
            }
            int i = service.updateItemBudget(id, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新預算資料成功，PROJECT_ID =" + Request["projectId"];
            }

            logger.Info("Request:PROJECT_ID =" + Request["projectId"]);
            return msg;
        }
        /// <summary>
        /// 異動單控制區塊
        /// </summary>
        /// <returns></returns>
        //異動單管理資料
        public ActionResult PlanItemChange()
        {
            string projectId = Request["projectid"];
            logger.Info("projectid=" + projectId);
            TND_PROJECT p = service.getProject(projectId);
            ViewBag.projectId = p.PROJECT_ID;
            ViewBag.projectName = p.PROJECT_NAME;
            string status = null;
            string remark = null;
            //查詢條件
            if (null != Request["status"] && "*" != Request["status"])
            {
                logger.Debug("status=" + Request["status"]);
                status = Request["status"];
            }
            if (null != Request["remark"])
            {
                logger.Debug("remark=" + Request["remark"]);
                remark = Request["remark"];
            }

            CostChangeService cs = new CostChangeService();
            List<PLAN_COSTCHANGE_FORM> lstForms = cs.getChangeOrders(projectId, remark, status);
            return View(lstForms);
        }
        //建立異動單-標單品項選擇畫面
        public ActionResult createChangeForm()
        {
            logger.Debug("Action:" + Request["action"]);
            string id = Request["projectId"];
            ViewBag.projectId = id;
            //取得主系統資料
            Dictionary<string, object> sec = TypeSelectComponet.getMapItemQueryCriteria(id);
            ViewBag.SystemMain = sec["SystemMain"];
            ViewBag.SystemSub = sec["SystemSub"];
            ViewBag.TypeCodeL1 = sec["TypeCodeL1"];
            return View();
        }

        //取得異動單相關品項選項
        public ActionResult getMapItem4ChangeForm(FormCollection f)
        {
            string projectid, typeCode1, typeCode2, systemMain, systemSub, primeside, primesideName, secondside, secondsideName, mapno, buildno, devicename, mapType, strart_id, end_id;
            ProjectPlanService planService = new ProjectPlanService();
            TypeSelectComponet.getMapItem(f, out projectid, out typeCode1, out typeCode2, out systemMain, out systemSub, out primeside, out primesideName, out secondside, out secondsideName, out mapno, out buildno, out devicename, out mapType, out strart_id, out end_id);
            if (null == f["mapType"] || "" == f["mapType"])
            {
                ViewBag.Message = "至少需選擇一項施作項目!!";
                return PartialView("_getMapItem4ChangeForm", null);
            }
            string[] mapTypes = mapType.Split(',');
            for (int i = 0; i < mapTypes.Length; i++)
            {
                switch (mapTypes[i])
                {
                    case "MAP_DEVICE"://設備
                        logger.Debug("MapType: MAP_DEVICE(設備)");
                        //增加九宮格、次九宮格、主系統、次系統等條件
                        planService.getMapItem(projectid, devicename, strart_id, end_id, typeCode1, typeCode2, systemMain, systemSub);
                        break;
                    case "MAP_PEP"://電器管線
                        logger.Debug("MapType: MAP_PEP(電器管線)");
                        //增加一次側名稱、二次側名稱
                        planService.getMapPEP(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        break;
                    case "MAP_LCP"://弱電管線
                        logger.Debug("MapType: MAP_LCP(弱電管線)");
                        planService.getMapLCP(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        break;
                    case "TND_MAP_PLU"://給排水
                        logger.Debug("MapType: TND_MAP_PLU(給排水)");
                        planService.getMapPLU(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        break;
                    case "MAP_FP"://消防電
                        logger.Debug("MapType: MAP_FP(消防電)");
                        planService.getMapFP(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        break;
                    case "MAP_FW"://消防水
                        planService.getMapFW(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        logger.Debug("MapType: MAP_FW(消防水)");
                        break;
                    default:
                        logger.Debug("MapType nothing!!");
                        break;
                }
            }
            ViewBag.Message = planService.resultMessage;
            return PartialView("_getMapItem4ChangeForm", planService.viewModel);
        }
        //成本異動單表單
        public ActionResult costChangeForm(string id)
        {
            string formId = id;
            logger.Debug("formId=" + formId);
            //fm3
            CostChangeService cs = new CostChangeService();
            cs.getChangeOrderForm(formId);
            ViewBag.FormId = formId;
            ViewBag.Remark = cs.form.REMARK;
            ViewBag.projectId = cs.project.PROJECT_ID;
            ViewBag.projectName = cs.project.PROJECT_NAME;
            ViewBag.formStatus = cs.form.STATUS;
            ViewBag.settlementDate = cs.form.SETTLEMENT_DATE;
            return View(cs.lstItem);
        }
        //建立與修改異動單--加入審核功能
        public string creatOrModifyChangeForm(FormCollection f)
        {
            SYS_USER u = (SYS_USER)Session["user"];
            string formId = f["txtFormId"].Trim();
            string projectId = f["projectId"].Trim();
            string remark = null;
            if (null != f["remark"])
            {
                remark = f["remark"].Trim();
            }

            logger.Debug("projectId=" + projectId + ",formID=" + formId + ",remak=" + remark);
            string reurnMsg = "";

            PLAN_COSTCHANGE_FORM formCostChange = new PLAN_COSTCHANGE_FORM();
            List<PLAN_COSTCHANGE_ITEM> lstItemId = new List<PLAN_COSTCHANGE_ITEM>();

            if (formId != "")
            {
                //直接新增時使用者尚未建立明細
                logger.Debug("Modify Change Order:" + formId);
                formCostChange.FORM_ID = formId;
                formCostChange.REMARK = remark;
                formCostChange.STATUS = f["status"];
                if (null != f["status_next"])
                {
                    //審核通過或刪除
                    formCostChange.STATUS = f["status_next"];
                    formCostChange.SETTLEMENT_DATE = DateTime.Parse(f["settlementDate"]);
                }
                formCostChange.MODIFY_DATE = DateTime.Now;
                formCostChange.MODIFY_USER_ID = u.USER_ID;
                logger.Debug("Item Id=" + f["uid"] + "," + f["itemdesc"]);
                if (null != f["uid"])
                {
                    string[] itemId = f["uid"].Split(',');
                    string[] itemdesc = f["itemdesc"].Split(',');
                    string[] itemunit = f["itemunit"].Split(',');
                    string[] itemUnitPrice = f["itemUnitPrice"].Split(',');
                    string[] itemQty = f["itemQty"].Split(',');
                    string[] itemRemark = f["itemRemark"].Split(',');

                    for (int i = 0; i < itemId.Length; i++)
                    {
                        PLAN_COSTCHANGE_ITEM it = new PLAN_COSTCHANGE_ITEM();
                        it.ITEM_UID = long.Parse(itemId[i]);
                        it.ITEM_DESC = itemdesc[i];
                        it.ITEM_UNIT = itemunit[i];
                        if (itemUnitPrice[i] != "")
                        {
                            it.ITEM_UNIT_PRICE = decimal.Parse(itemUnitPrice[i]);
                        }
                        if (itemQty[i] != "")
                        {
                            it.ITEM_QUANTITY = decimal.Parse(itemQty[i]);
                        }
                        it.ITEM_REMARK = itemRemark[i];
                        if (null != f["transFlg." + itemId[i]])
                        {
                            it.TRANSFLAG = f["transFlg." + itemId[i]];
                        }
                        else
                        {
                            it.TRANSFLAG = "0";
                        }
                        logger.Debug(it.ITEM_UID + "," + it.ITEM_DESC + "," + it.ITEM_UNIT_PRICE + "," + it.ITEM_QUANTITY + "," + it.ITEM_REMARK + "," + it.TRANSFLAG);
                        it.MODIFY_DATE = DateTime.Now;
                        it.MODIFY_USER_ID = u.USER_ID;
                        lstItemId.Add(it);
                    }
                }
                CostChangeService s = new CostChangeService();
                reurnMsg = s.updateChangeOrder(formCostChange, lstItemId) + " <a href='/Plan/costChangeForm/" + formId + "'>返回</a>";
            }
            else
            {
                //新增異動單
                formCostChange.PROJECT_ID = projectId;
                formCostChange.REMARK = remark;
                formCostChange.CREATE_DATE = DateTime.Now;
                formCostChange.CREATE_USER_ID = u.USER_ID;
                //設備
                logger.Debug("MapType: MAP_DEVICE(設備):" + f["map_device"]);
                if (null != f["map_device"])
                {
                    lstItemId.AddRange(getItem("map_device.", f["map_device"].Trim().Split(',')));
                }
                //電器管線
                logger.Debug("MapType: MAP_PEP(電器管線):" + f["map_pep"]);
                if (null != f["map_pep"])
                {
                    lstItemId.AddRange(getItem("map_pep.", f["map_pep"].Trim().Split(',')));
                }
                //弱電管線
                logger.Debug("MapType: MAP_LCP(弱電管線)+" + f["map_lcp"]);
                if (null != f["map_lcp"])
                {
                    lstItemId.AddRange(getItem("map_lcp.", f["map_lcp"].Trim().Split(',')));
                }
                //給排水
                logger.Debug("MapType: TND_MAP_PLU(給排水):" + f["map_plu"]);
                if (null != f["map_plu"])
                {
                    lstItemId.AddRange(getItem("map_plu.", f["map_plu"].Trim().Split(',')));
                }
                //消防電
                logger.Debug("MapType: MAP_FP(消防電):" + f["map_fp"]);
                if (null != f["map_fp"])
                {
                    lstItemId.AddRange(getItem("map_fp.", f["map_fp"].Trim().Split(',')));
                }
                //消防水
                logger.Debug("MapType: MAP_FW(消防水):" + f["map_fw"]);
                if (null != f["map_fw"])
                {
                    lstItemId.AddRange(getItem("map_fw.", f["map_fw"].Trim().Split(',')));
                }
                CostChangeService s = new CostChangeService();
                reurnMsg = s.createChangeOrder(formCostChange, lstItemId);
            }
            return reurnMsg;
        }
        //資料轉換
        private IEnumerable<PLAN_COSTCHANGE_ITEM> getItem(string prefix, string[] aryPlanItemIds)
        {
            List<PLAN_COSTCHANGE_ITEM> lstItemId = new List<PLAN_COSTCHANGE_ITEM>();
            for (int i = 0; i < aryPlanItemIds.Length; i++)
            {
                PLAN_COSTCHANGE_ITEM item = new PLAN_COSTCHANGE_ITEM();
                item.PLAN_ITEM_ID = aryPlanItemIds[i];
                logger.Debug(item.ITEM_ID + " Qty = " + Request[prefix + item.ITEM_ID]);
                if (Request[prefix + item.PLAN_ITEM_ID].Trim() != "")
                {
                    item.ITEM_QUANTITY = int.Parse(Request[prefix + item.PLAN_ITEM_ID]);
                }else
                {
                    item.ITEM_QUANTITY = 0;
                }
                logger.Debug("Item_ID=" + item.ITEM_ID + ",Change Qty=" + item.ITEM_QUANTITY);
                lstItemId.Add(item);
            }
            return lstItemId;
        }
        //新增異動單品項
        public String addChangeOrderItem(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "更新成功!!";

            PLAN_COSTCHANGE_ITEM item = new PLAN_COSTCHANGE_ITEM();
            item.FORM_ID = form["dia_form_id"];
            item.PLAN_ITEM_ID = form["dia_plan_item_id"];
            item.ITEM_ID = form["item_id"];
            item.ITEM_DESC = form["dia_item_desc"];
            item.ITEM_UNIT = form["dia_item_unit"];
            try
            {
                item.ITEM_QUANTITY = decimal.Parse(form["dia_item_quantity"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not quattity:" + ex.Message);
            }
            try
            {
                item.ITEM_UNIT_PRICE = decimal.Parse(form["dia_item_unit_price"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not unit price:" + ex.Message);
            }
            item.ITEM_REMARK = form["dia_item_remark"];

            item.TRANSFLAG = form["dia_transFlag"];

            SYS_USER loginUser = (SYS_USER)Session["user"];
            item.CREATE_USER_ID = loginUser.USER_ID;
            item.CREATE_DATE = DateTime.Now;
            // InquiryFormService service = new InquiryFormService();

            CostChangeService cs = new CostChangeService();
            int i = cs.addChangeOrderItem(item);

            //  if (i == 0) { msg = service.message; }
            return msg + "(" + i + ")";
        }
        //刪除單一品項資料
        public String delChangeOrderItem()
        {
            long itemUid = long.Parse(Request["itemid"]);
            SYS_USER loginUser = (SYS_USER)Session["user"];
            logger.Info(loginUser.USER_ID + " remove data:change_order_item uid=" + itemUid);
            CostChangeService cs = new CostChangeService();
            int i = cs.delChangeOrderItem(itemUid);
            return "資料已刪除(" + i + ")";
        }
        //將異動單送審
        public string send2Audit()
        {
            string strFormID = null;
            string returnMsg = null;
            SYS_USER loginUser = (SYS_USER)Session["user"];
            if (null != Request["formId"])
            {
                strFormID = Request["formId"];
                logger.Info(loginUser.USER_ID + " set form to audit:" + strFormID);

                PLAN_COSTCHANGE_FORM formCostChange = new PLAN_COSTCHANGE_FORM();
                formCostChange.FORM_ID = strFormID;
                formCostChange.REMARK = ""; //如果需要可透過Remark 補充資料
                formCostChange.STATUS = "送審";
                formCostChange.MODIFY_DATE = DateTime.Now;
                formCostChange.MODIFY_USER_ID = loginUser.USER_ID;
                CostChangeService cs = new CostChangeService();
                returnMsg = cs.updateChangeOrderStatus(formCostChange);
            }
            return returnMsg;
        }
        //下載異動單
        public void downloadCostChangeForm()
        {
            string formid = Request["formId"];
            logger.Debug("Download costchange form=" + formid);
            CostChangeService chService = new CostChangeService();
            chService.getChangeOrderForm(formid);
            poi4CostChangeService poiservice = new poi4CostChangeService();
            poiservice.createExcel(chService.project, chService.form, chService.lstItem);
            string fileLocation = poiservice.outputFile;

            //檔案名稱 HttpUtility.UrlEncode預設會以UTF8的編碼系統進行QP(Quoted-Printable)編碼，可以直接顯示的7 Bit字元(ASCII)就不用特別轉換。
            string filename = HttpUtility.UrlEncode(Path.GetFileName(fileLocation));
            Response.Clear();
            Response.Charset = "utf-8";
            Response.ContentType = "text/xls";
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", filename));
            Response.WriteFile(fileLocation);
            Response.End();
        }
        //成本預算管制表
        public ActionResult costControllerIndex(string id)
        {
            string projectId = id;// Request["id"];
            ContextService4PlanCost costService = new ContextService4PlanCost();
            //成本預算管制表物件
            costService.getCostControlInfo(projectId);
            return View(costService.CostInfo);
        }
        //建立間接成本
        public string createIndirectCost()
        {
            string projectId = Request["projectId"];
            SYS_USER u = (SYS_USER)Session["user"];
            ContextService4PlanCost s = new ContextService4PlanCost();
            s.createIndirectCost(projectId, u.USER_ID);
            logger.Debug("create indirect cost by projectid=" + projectId);
            return "建立成功";
        }
        //更新間接成本\
        public string modifyIndirectCost()
        {
            string projectId = Request["projectId"];
            string[] fieldIds = Request["fieldId"].Split(',');
            string[] costs = Request["cost"].Split(',');
            string[] notes = Request["note"].Split(',');
            logger.Debug("Field Count=" + fieldIds.Count() + ",Cost Count=" + costs.Count());
            SYS_USER u = (SYS_USER)Session["user"];
            ContextService4PlanCost s = new ContextService4PlanCost();
            List<PLAN_INDIRECT_COST> items = new List<PLAN_INDIRECT_COST>();
            for (int i = 0; i < fieldIds.Count(); i++)
            {
                PLAN_INDIRECT_COST it = new PLAN_INDIRECT_COST();
                it.FIELD_ID = fieldIds[i];
                it.COST = decimal.Parse(costs[i]);
                it.NOTE = notes[i];
                it.MODIFY_ID = u.USER_ID;
                items.Add(it);
            }
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(items);
            logger.Debug("item  info=" + itemJson);
            s.modifyIndirectCost(projectId, items);
            //logger.Debug("create indirect cost by projectid=" + projectId);
            return "修改成功";
        }
        //上傳異動單
        public string uploadCostChangeForm(HttpPostedFileBase file1)
        {
            string projectid = Request["projectid"];
            string fromid = Request["formid"];
            logger.Debug("ProjectID=" + projectid + ",Upload ProjectItem=" + file1.FileName);
            SYS_USER u = (SYS_USER)Session["user"];

            if (null != file1 && file1.ContentLength != 0)
            {
                try
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + file1.FileName);
                    //2.1 檢查目錄
                    var fileName = Path.GetFileName(file1.FileName);
                    string folder = ContextService.strUploadPath + "/" + projectid + "/CostChangeForm";
                    ZipFileCreator.CreateDirectory(folder);
                    logger.Debug("costchange form folder");
                    //2.1 將上傳檔案存檔
                    var path = Path.Combine(folder, fileName);
                    logger.Info("save excel file:" + path);
                    file1.SaveAs(path);
                    //2.2 解析Excel 檔案
                    poi4CostChangeService poiService = new poi4CostChangeService();
                    poiService.setUser(u);
                    poiService.getDataFromExcel(path, projectid, fromid);
                    //System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                    //string itemJson = objSerializer.Serialize(poiService.project);
                    //logger.Info("project info=" + itemJson);
                    //itemJson = objSerializer.Serialize(poiService.costChangeForm);
                    //logger.Info("form info=" + itemJson);
                    //itemJson = objSerializer.Serialize(poiService.lstItem);
                    //logger.Info("item info=" + itemJson);
                    //2.3 寫入資料
                    CostChangeService s = new CostChangeService();
                    if (null== poiService.costChangeForm.FORM_ID || poiService.costChangeForm.FORM_ID == "")
                    {
                        s.createChangeOrder(poiService.costChangeForm, poiService.lstItem);
                    }
                    else
                    {
                        s.updateChangeOrder(poiService.costChangeForm, poiService.lstItem);
                    }

                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
                    return ex.Message;
                }
            }
            return "匯入成功!!";

        }
    }
}

